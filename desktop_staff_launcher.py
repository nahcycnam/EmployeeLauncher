import sys
import os
import json
import subprocess
import ctypes
from ctypes import wintypes
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTableWidget,
                             QTableWidgetItem, QVBoxLayout, QWidget)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor, QPaintEvent, QFont
from PyQt6.QtNetwork import QLocalServer, QLocalSocket


# --- Windows API 常量与函数定义 ---
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP   = 0x0205
SW_SHOW = 5

user32 = ctypes.windll.user32

# 定义 POINT 结构
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

# 获取桌面窗口句柄（适用于 Windows 10/11）
def get_desktop_hwnd():
    # 首先尝试找到 SysListView32 (桌面图标列表)
    def find_syslistview_hwnd(hwnd, lparam):
        class_name = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_name, 256)
        if class_name.value == "SysListView32":
            # 找到直接返回，通过 ctypes 的返回值无法直接传递，使用全局变量
            found_hwnd = ctypes.c_void_p()
            found_hwnd.value = hwnd
            # 将 hwnd 存入 lparam 指向的指针
            ctypes.memmove(lparam, ctypes.byref(found_hwnd), ctypes.sizeof(ctypes.c_void_p))
            return False  # 继续枚举，但实际已经记录
        return True

    # 先找 WorkerW 或 Progman
    desktop = user32.FindWindowW("Progman", None)
    if not desktop:
        desktop = user32.FindWindowW("WorkerW", None)
    if not desktop:
        return None

    # 枚举子窗口查找 SysListView32
    # 使用 EnumChildWindows 回调
    found = wintypes.HWND()
    def enum_callback(hwnd, lparam):
        class_name = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_name, 256)
        if class_name.value == "SysListView32":
            # 找到，保存到 lparam 指向的 HWND
            ctypes.memmove(lparam, ctypes.byref(wintypes.HWND(hwnd)), ctypes.sizeof(wintypes.HWND))
            return False  # 停止枚举
        return True

    EnumChildWindows = user32.EnumChildWindows
    EnumChildWindows.argtypes = [wintypes.HWND, ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
    EnumChildWindows.restype = ctypes.c_bool

    # 定义回调
    callback_type = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    callback = callback_type(enum_callback)

    result_hwnd = wintypes.HWND()
    EnumChildWindows(desktop, callback, wintypes.LPARAM(ctypes.addressof(result_hwnd)))
    if result_hwnd.value:
        return result_hwnd.value
    # 如果没找到，直接返回桌面顶层窗口
    return desktop


class DesktopTableWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 加载 JSON 配置
        self.config = self.load_config()

        # 设置无边框、置底、透明背景
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool
        )

        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)

        self.setStyleSheet("""
            QMainWindow {
                background-color: transparent;
                border: none;
                margin: 0px;
                padding: 0px;
            }
        """)

        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(0, 0, 0, 0))
        self.setPalette(palette)

        central_widget = QWidget()
        central_widget.setStyleSheet("""
            QWidget {
                background-color: transparent;
                border: none;
                margin: 0px;
                padding: 0px;
            }
        """)
        central_widget.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.table = ClickableTableWidget()
        self.table.parent_window = self  # 让表格能访问主窗口

        # 应用 JSON 中的字体设置（QFont）
        font_cfg = self.config["table"]
        font = QFont()
        font.setFamily(font_cfg["font_family"].split(",")[0].strip())
        font.setPointSize(font_cfg["font_size"])
        font.setBold(font_cfg["font_weight"] == "bold")
        font.setItalic(font_cfg["italic"])
        self.table.setFont(font)

        # 动态生成样式表
        style_sheet = f"""
            QTableWidget {{
                background-color: transparent;
                border: none;
                outline: 0;
                margin: 0px;
                padding: 0px;
                font-family: {font_cfg["font_family"]};
                font-size: {font_cfg["font_size"]}px;
                font-weight: {font_cfg["font_weight"]};
                color: {font_cfg["color"]};
            }}
            QTableWidget::item {{
                background-color: transparent;
                border: none;
                margin: 0px;
                padding: {font_cfg["item_padding"]};
            }}
            QTableWidget::item:selected {{
                background-color: transparent;
                color: #FF0000;
            }}
            QHeaderView::section {{
                background-color: transparent;
                border: none;
                margin: 0px;
                padding: 0px;
            }}
            QTableCornerButton::section {{
                background-color: transparent;
                border: none;
                margin: 0px;
                padding: 0px;
            }}
        """
        self.table.setStyleSheet(style_sheet)

        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setShowGrid(False)

        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setVisible(False)
        self.table.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        layout.addWidget(self.table)

        self.load_excel_data()

        # 窗口坐标从 JSON 读取
        self.window_x = self.config["window"]["x"]
        self.window_y = self.config["window"]["y"]

        self.adjust_window_size()
        self.set_window_to_bottom()

    def load_config(self):
        default_config = {
            "window": {"x": 1343, "y": 15},
            "table": {
                "font_family": "Microsoft YaHei, 微软雅黑, Segoe UI",
                "font_size": 22,
                "font_weight": "normal",
                "italic": False,
                "color": "#FFFFFF",
                "item_padding": "2px 5px"
            }
        }
        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
                for k, v in default_config.items():
                    if k not in config:
                        config[k] = v
                    elif isinstance(v, dict):
                        for sub_k, sub_v in v.items():
                            if sub_k not in config.get(k, {}):
                                config.setdefault(k, {})[sub_k] = sub_v
                return config
        except FileNotFoundError:
            print(f"警告：找不到 {config_path}，使用默认配置")
            return default_config
        except Exception as e:
            print(f"警告：读取配置文件出错 {e}，使用默认配置")
            return default_config

    def load_excel_data(self):
        try:
            df = pd.read_excel("people.xlsx")
            names = df.iloc[:, 0].dropna().tolist()
            job_numbers = df.iloc[:, 1].dropna().tolist()

            self.table.setRowCount(len(names))
            self.table.setColumnCount(2)

            for row, (name, job_num) in enumerate(zip(names, job_numbers)):
                name_item = QTableWidgetItem(str(name))
                job_item = QTableWidgetItem(str(job_num))
                name_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                job_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

                self.table.setItem(row, 0, name_item)
                self.table.setItem(row, 1, job_item)

            self.table.resizeColumnsToContents()
            self.table.resizeRowsToContents()

        except FileNotFoundError:
            print("错误：找不到 people.xlsx 文件")
        except Exception as e:
            print(f"读取文件时出错：{e}")

    def adjust_window_size(self):
        width = sum(self.table.columnWidth(i) for i in range(self.table.columnCount()))
        height = sum(self.table.rowHeight(i) for i in range(self.table.rowCount()))
        self.resize(width, height)
        self.move(self.window_x, self.window_y)

    def on_cell_double_click(self, row, column):
        if column == 0:
            name = self.table.item(row, 0).text()
            print(f"双击了姓名：{name}")
            if name == "吴文超":
                exe_path = r"C:\Program Files (x86)\广州地铁\Glink\sbdt\shigongW.exe"
                self.run_external_program(exe_path)
            elif name == "陈孟熙":
                exe_path = r"E:\陈孟熙\yp\Project\施工系统\施工登录\受令当班登记.exe"
                self.run_external_program(exe_path)
            else:
                print(f"未配置姓名“{name}”对应的程序")
        elif column == 1:
            job_num_str = self.table.item(row, 1).text()
            try:
                job_num = int(job_num_str)
            except ValueError:
                print(f"工号格式错误：{job_num_str}")
                return
            print(f"双击了工号：{job_num}")
            if job_num in (116873, 120557):
                exe_path = r"E:\陈孟熙\yp\Project\运营前检查\github\运营前检查.exe"
                self.run_external_program(exe_path)
            else:
                print(f"未配置工号 {job_num} 对应的程序")

    def run_external_program(self, exe_path):
        """运行外部程序，并处理可能的异常"""
        try:
            subprocess.Popen([exe_path], shell=True)
            print(f"已启动：{exe_path}")
        except FileNotFoundError:
            print(f"错误：找不到文件 {exe_path}")
        except Exception as e:
            print(f"启动程序时出错：{e}")

    def send_right_click_to_desktop(self, global_x, global_y):
        """向桌面窗口发送右键消息，弹出原生右键菜单"""
        desktop_hwnd = get_desktop_hwnd()
        if not desktop_hwnd:
            print("警告：未找到桌面窗口句柄")
            return

        # 将屏幕坐标转换为桌面窗口客户区坐标
        pt = POINT(global_x, global_y)
        user32.ScreenToClient(desktop_hwnd, ctypes.byref(pt))

        # 发送鼠标按下和弹起消息
        lparam = (pt.y << 16) | (pt.x & 0xFFFF)
        user32.PostMessageW(desktop_hwnd, WM_RBUTTONDOWN, 0, lparam)
        user32.PostMessageW(desktop_hwnd, WM_RBUTTONUP, 0, lparam)
        print(f"已向桌面发送右键消息，坐标：({global_x}, {global_y})")

    def paintEvent(self, event: QPaintEvent):
        super().paintEvent(event)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == event.Type.ActivationChange:
            self.set_window_to_bottom()

    def set_window_to_bottom(self):
        try:
            HWND_BOTTOM = 1
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_NOACTIVATE = 0x0010
            SWP_SHOWWINDOW = 0x0040
            hwnd = int(self.winId())
            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW
            )
        except Exception as e:
            print(f"设置置底窗口时出错：{e}")

    def showEvent(self, event):
        super().showEvent(event)
        self.set_window_to_bottom()

    def activate_window(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()


class ClickableTableWidget(QTableWidget):
    def mouseDoubleClickEvent(self, event):
        pos = self.viewport().mapFrom(self, event.position().toPoint())
        row = self.rowAt(pos.y())
        col = self.columnAt(pos.x())
        print(f"双击: pos={pos}, row={row}, col={col}")

        if row >= 0 and col >= 0:
            parent = getattr(self, 'parent_window', None)
            if parent is None:
                parent = self.window()
            if hasattr(parent, 'on_cell_double_click'):
                parent.on_cell_double_click(row, col)
            else:
                print("父窗口没有 on_cell_double_click 方法")
        else:
            print("未命中有效单元格")

    def contextMenuEvent(self, event):
        """右键事件：穿透给真正的桌面"""
        # 获取鼠标的全局坐标
        global_pos = event.globalPos()
        parent = getattr(self, 'parent_window', None)
        if parent is None:
            parent = self.window()
        if hasattr(parent, 'send_right_click_to_desktop'):
            parent.send_right_click_to_desktop(global_pos.x(), global_pos.y())
        else:
            print("父窗口没有 send_right_click_to_desktop 方法")
        event.accept()  # 阻止表格本身弹出任何菜单，但事件已转发


class SingleInstance:
    SERVER_NAME = "DesktopTableAppSingleton"

    def __init__(self, app, create_window_callback):
        self.app = app
        self.create_window_callback = create_window_callback
        self.server = None
        self.window = None

    def try_run(self):
        socket = QLocalSocket()
        socket.connectToServer(self.SERVER_NAME)
        if socket.waitForConnected(1000):
            socket.write(b"activate")
            socket.flush()
            socket.waitForBytesWritten(1000)
            socket.disconnectFromServer()
            print("程序已在运行，已发送激活信号，本实例退出。")
            return False
        else:
            self.server = QLocalServer()
            self.server.listen(self.SERVER_NAME)
            self.server.newConnection.connect(self.handle_new_connection)
            self.window = self.create_window_callback()
            self.window.show()
            return True

    def handle_new_connection(self):
        if self.server.hasPendingConnections():
            conn = self.server.nextPendingConnection()
            conn.readyRead.connect(lambda: self.process_message(conn))
            self.activate_existing_window()

    def process_message(self, conn):
        data = conn.readAll().data()
        if data == b"activate":
            self.activate_existing_window()
        conn.disconnectFromServer()

    def activate_existing_window(self):
        if self.window is not None:
            print("激活现有窗口...")
            self.window.activate_window()


def create_window():
    return DesktopTableWindow()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    palette = app.palette()
    palette.setColor(QPalette.ColorRole.Window, QColor(0, 0, 0, 0))
    palette.setColor(QPalette.ColorRole.Base, QColor(0, 0, 0, 0))
    app.setPalette(palette)

    single = SingleInstance(app, create_window)
    if not single.try_run():
        sys.exit(0)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
