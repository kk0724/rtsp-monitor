import sys
import cv2
import win32gui
import win32con
import numpy as np
import winsound
from datetime import datetime
from PyQt5.QtGui import QImage, QPixmap, QIcon
from PyQt5.QtCore import QTimer, Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QVBoxLayout, 
                             QSystemTrayIcon, QStyle, QMenu, QAction, 
                             QDialog, QLineEdit, QPushButton, QHBoxLayout, 
                             QMessageBox, QGroupBox, QFormLayout)
import time
import queue
import os
import json


class MagnifyWindow(QWidget):
    """放大显示窗口"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("运动目标放大")
        self.setFixedSize(400, 300)  # 增大窗口尺寸
        self.setWindowFlag(Qt.WindowStaysOnTopHint)  # 始终置顶
        self.setWindowFlag(Qt.FramelessWindowHint)  # 无边框，更简洁
        self.setAttribute(Qt.WA_TranslucentBackground)  # 透明背景
        
        # 设置窗口位置在原监控画面的正上方
        if parent:
            parent_geo = parent.geometry()
            # 计算居中位置
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() - self.height() - 10  # 上方10像素
            
            # 确保不超出屏幕顶部
            if y < 0:
                y = parent_geo.y() + parent_geo.height() + 10  # 如果上方放不下，放到下方
            
            self.move(x, y)
            print(f"放大窗口位置: ({x}, {y})")
        
        # 创建一个容器来放内容，带边框和阴影效果
        self.container = QLabel(self)
        self.container.setGeometry(5, 5, self.width()-10, self.height()-10)
        self.container.setStyleSheet("""
            background-color: black;
            border: 3px solid red;
            border-radius: 5px;
        """)
        
        self.label = QLabel(self.container)
        self.label.setGeometry(5, 5, self.container.width()-10, self.container.height()-10)
        self.label.setScaledContents(True)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("background-color: black;")
        
        # 标题栏（用于拖动）
        self.title_bar = QLabel(self)
        self.title_bar.setGeometry(5, 5, self.width()-10, 25)
        self.title_bar.setStyleSheet("""
            background-color: rgba(255, 0, 0, 180);
            color: white;
            font-weight: bold;
            padding-left: 5px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
        """)
        self.title_bar.setText(" 运动目标放大")
        self.title_bar.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.title_bar.hide()  # 默认隐藏，鼠标移入时显示
        
        # 关闭按钮
        self.close_btn = QPushButton("×", self)
        self.close_btn.setGeometry(self.width()-30, 5, 25, 25)
        self.close_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 0, 0, 180);
                color: white;
                font-weight: bold;
                font-size: 16px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: rgba(255, 50, 50, 200);
            }
        """)
        self.close_btn.clicked.connect(self.hide)
        self.close_btn.hide()
        
        # 当前放大的区域
        self.current_roi = None
        self.hide_timer = QTimer(self)
        self.hide_timer.timeout.connect(self.hide)
        self.hide_timer.setSingleShot(True)
        
        # 鼠标跟踪
        self.setMouseTracking(True)
        self.container.setMouseTracking(True)
        self.label.setMouseTracking(True)
        self.title_bar.setMouseTracking(True)
        
        # 拖动相关
        self.dragging = False
        self.drag_position = None
        
        # 先隐藏，等有运动再显示
        self.hide()
        print(f"{datetime.now()} 放大窗口已创建")
    
    def enterEvent(self, event):
        """鼠标进入窗口时显示标题栏和关闭按钮"""
        self.title_bar.show()
        self.close_btn.show()
        super().enterEvent(event)
    
    def leaveEvent(self, event):
        """鼠标离开窗口时隐藏标题栏和关闭按钮"""
        if not self.dragging:  # 拖动时不隐藏
            self.title_bar.hide()
            self.close_btn.hide()
        super().leaveEvent(event)
    
    def mousePressEvent(self, event):
        """点击窗口可以拖动"""
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
    
    def mouseMoveEvent(self, event):
        """允许拖动窗口"""
        if event.buttons() == Qt.LeftButton and self.dragging:
            self.move(event.globalPos() - self.drag_position)
            event.accept()
    
    def mouseReleaseEvent(self, event):
        """释放鼠标结束拖动"""
        if event.button() == Qt.LeftButton:
            self.dragging = False
            # 如果不是在窗口内，隐藏标题栏
            if not self.underMouse():
                self.title_bar.hide()
                self.close_btn.hide()
            event.accept()
    
    def mouseDoubleClickEvent(self, event):
        """双击切换置顶状态"""
        if event.button() == Qt.LeftButton:
            if self.windowFlags() & Qt.WindowStaysOnTopHint:
                self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
                self.title_bar.setStyleSheet("""
                    background-color: rgba(100, 100, 100, 180);
                    color: white;
                    font-weight: bold;
                    padding-left: 5px;
                    border-top-left-radius: 5px;
                    border-top-right-radius: 5px;
                """)
                self.close_btn.setStyleSheet("""
                    QPushButton {
                        background-color: rgba(100, 100, 100, 180);
                        color: white;
                        font-weight: bold;
                        font-size: 16px;
                        border: none;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: rgba(150, 150, 150, 200);
                    }
                """)
            else:
                self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
                self.title_bar.setStyleSheet("""
                    background-color: rgba(255, 0, 0, 180);
                    color: white;
                    font-weight: bold;
                    padding-left: 5px;
                    border-top-left-radius: 5px;
                    border-top-right-radius: 5px;
                """)
                self.close_btn.setStyleSheet("""
                    QPushButton {
                        background-color: rgba(255, 0, 0, 180);
                        color: white;
                        font-weight: bold;
                        font-size: 16px;
                        border: none;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: rgba(255, 50, 50, 200);
                    }
                """)
            self.show()
    
    def showEvent(self, event):
        """窗口显示时触发"""
        super().showEvent(event)
        self.raise_()
        self.activateWindow()
        print(f"{datetime.now()} 放大窗口已显示")
    
    def update_magnified_view(self, frame, roi_rect):
        """
        更新放大视图
        roi_rect: (x, y, w, h) 在原图中的坐标和大小
        """
        if frame is None or roi_rect is None:
            return
        
        x, y, w, h = roi_rect
        
        # 确保ROI在图像范围内
        x = max(0, min(x, frame.shape[1] - 1))
        y = max(0, min(y, frame.shape[0] - 1))
        w = min(w, frame.shape[1] - x)
        h = min(h, frame.shape[0] - y)
        
        if w <= 0 or h <= 0:
            return
        
        # 提取ROI区域
        roi_frame = frame[y:y+h, x:x+w]
        
        if roi_frame.size > 0:
            # 获取显示区域大小
            display_width = self.label.width()
            display_height = self.label.height()
            
            if display_width <= 0 or display_height <= 0:
                display_width = 380
                display_height = 280
            
            # 获取ROI的宽高比
            roi_aspect = w / h
            display_aspect = display_width / display_height
            
            if roi_aspect > display_aspect:
                new_w = display_width
                new_h = int(display_width / roi_aspect)
            else:
                new_h = display_height
                new_w = int(display_height * roi_aspect)
            
            try:
                magnified = cv2.resize(roi_frame, (new_w, new_h))
            except Exception as e:
                print(f"缩放失败: {e}")
                return
            
            # 创建黑色背景的画布
            canvas = np.zeros((display_height, display_width, 3), dtype=np.uint8)
            
            # 将放大后的图像居中放置
            y_offset = (display_height - new_h) // 2
            x_offset = (display_width - new_w) // 2
            canvas[y_offset:y_offset+new_h, x_offset:x_offset+new_w] = magnified
            
            # 添加信息
            cv2.putText(canvas, f"Size: {w}x{h}", (10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
            
            # 添加时间戳
            time_str = datetime.now().strftime("%H:%M:%S")
            cv2.putText(canvas, time_str, (display_width-100, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
            
            # 转换并显示
            rgb = cv2.cvtColor(canvas, cv2.COLOR_BGR2RGB)
            h, w, ch = rgb.shape
            bytes_per_line = ch * w
            qimg = QImage(rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.label.setPixmap(pixmap)
            
            self.current_roi = roi_rect
            
            # 确保窗口显示并置顶
            if self.isHidden():
                self.show()
            else:
                self.raise_()
                self.activateWindow()
            
            # 重置隐藏定时器（5秒后自动隐藏）
            self.hide_timer.stop()
            self.hide_timer.start(5000)
    
    def closeEvent(self, event):
        """关闭窗口时隐藏而不是销毁"""
        event.ignore()
        self.hide()
        self.hide_timer.stop()


class VideoCaptureThread(QThread):
    frame_ready = pyqtSignal(object)
    
    def __init__(self, camera_url):
        super().__init__()
        self.camera_url = camera_url
        self.running = True
        self.cap = None
        
    def run(self):
        os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = "rtsp_transport;tcp;buffer_size=1024000"
        
        self.cap = cv2.VideoCapture(self.camera_url, cv2.CAP_FFMPEG)
        self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        self.cap.set(cv2.CAP_PROP_FPS, 30)
        
        time.sleep(1)
        
        print(f"{datetime.now()} 视频线程启动")
        
        while self.running:
            try:
                if self.cap is None or not self.cap.isOpened():
                    print(f"{datetime.now()} 重新连接摄像头...")
                    time.sleep(1)
                    self.cap = cv2.VideoCapture(self.camera_url, cv2.CAP_FFMPEG)
                    self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
                    continue
                
                ret, frame = self.cap.read()
                if ret and frame is not None:
                    self.frame_ready.emit(frame.copy())
                else:
                    print(f"{datetime.now()} 读取失败，准备重新连接...")
                    if self.cap:
                        self.cap.release()
                        self.cap = None
                    time.sleep(1)
                    
            except Exception as e:
                print(f"捕获线程异常: {e}")
                if self.cap:
                    self.cap.release()
                    self.cap = None
                time.sleep(1)
        
        if self.cap:
            self.cap.release()
        print(f"{datetime.now()} 视频线程结束")
    
    def stop(self):
        self.running = False
        self.wait()


class SettingsDialog(QDialog):
    """设置对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置 - 隐藏窗口")
        self.setFixedSize(450, 250)
        self.parent = parent
        
        self.settings = self.load_settings()
        
        layout = QVBoxLayout()
        
        info_label = QLabel("设置要自动隐藏的窗口名称（支持模糊匹配）：")
        layout.addWidget(info_label)
        
        form_layout = QFormLayout()
        self.window_name_edit = QLineEdit()
        self.window_name_edit.setText(self.settings.get("window_name", ""))
        self.window_name_edit.setPlaceholderText("例如：无标题 - 记事本、微信、Chrome")
        form_layout.addRow("窗口名称:", self.window_name_edit)
        layout.addLayout(form_layout)
        
        hint_label = QLabel("提示：可以输入窗口标题的一部分，程序会自动查找匹配的窗口")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        hint_label.setWordWrap(True)
        layout.addWidget(hint_label)
        
        self.window_list_btn = QPushButton("查看当前打开的窗口")
        self.window_list_btn.clicked.connect(self.show_window_list)
        layout.addWidget(self.window_list_btn)
        
        test_btn = QPushButton("测试查找窗口")
        test_btn.clicked.connect(self.test_find_window)
        layout.addWidget(test_btn)
        
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self.save_settings)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def load_settings(self):
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载设置失败: {e}")
        return {"window_name": ""}
    
    def save_settings(self):
        try:
            settings = {
                "window_name": self.window_name_edit.text().strip()
            }
            with open("settings.json", "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(self, "成功", "设置已保存！")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{e}")
    
    def test_find_window(self):
        window_name = self.window_name_edit.text().strip()
        if not window_name:
            QMessageBox.warning(self, "提示", "请输入窗口名称")
            return
        
        hwnd = win32gui.FindWindow(None, window_name)
        if hwnd:
            window_text = win32gui.GetWindowText(hwnd)
            QMessageBox.information(self, "成功", f"找到精确匹配窗口：{window_text}")
            return
        
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and window_name.lower() in title.lower():
                    windows.append(f"{title} (句柄: {hwnd})")
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            msg = "找到以下匹配窗口：\n" + "\n".join(windows[:10])
            QMessageBox.information(self, "成功", msg)
        else:
            QMessageBox.warning(self, "失败", "未找到匹配的窗口")
    
    def show_window_list(self):
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    windows.append(f"{title}")
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            msg = "当前打开的窗口：\n" + "\n".join(sorted(windows)[:20])
            QMessageBox.information(self, "窗口列表", msg)
        else:
            QMessageBox.information(self, "窗口列表", "没有找到可见窗口")


class VideoPlayer(QWidget):
    def __init__(self):
        super().__init__()

        os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.png.warning=false'
        
        self.camera_url = 'rtsp://admin:guest123@192.168.6.105:554/Streaming/Channels/101'
        
        self.roi = {
            "x": 360,
            "y": 400,
            "w": 1170,
            "h": 1040
        }
        
        self.roi_width = self.roi["w"]
        self.roi_height = self.roi["h"]
        
        self.display_width = 640
        self.display_height = int(self.display_width * self.roi_height / self.roi_width)
        
        self.fps = 0
        self.fps_counter = 0
        self.fps_last_time = time.time()
        
        self.frame_queue = queue.Queue(maxsize=2)
        self.last_frame = None
        self.frame_count = 0
        
        self.last_motion_contours = []
        self.motion_history = []
        
        print(f"{datetime.now()} 初始化视频播放器...")
        print(f"摄像头地址: {self.camera_url}")
        
        self.capture_thread = VideoCaptureThread(self.camera_url)
        self.capture_thread.frame_ready.connect(self.on_frame_received)
        self.capture_thread.start()

        self.hide_windows = False
        self.hwnd = -1
        self.window_name = ""
        
        self.load_settings()

        self.resize(320, 180)
        self.move(1600, 750)
        self.setMinimumSize(320, 180)
        self.setWindowTitle('监控 - ROI模式')
        self.setWindowFlag(Qt.WindowStaysOnTopHint)

        self.label = QLabel(self)
        self.label.setScaledContents(True)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("background-color: black;")

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.label)
        self.setLayout(layout)

        self.setup_tray_icon()

        self.bg_subtractor = cv2.createBackgroundSubtractorMOG2(
            history=100,
            varThreshold=36,
            detectShadows=False
        )

        self.change_counter = 0
        self.min_consecutive_changes = 1   

        self.initializing = True
        self.init_frames = 0
        self.init_max_frames = 50

        self.display_timer = QTimer(self)
        self.display_timer.timeout.connect(self.update_display)
        self.display_timer.start(20)

        self.show()
        
        self.change_detected_timer = QTimer(self)
        self.change_detected_timer.timeout.connect(self.reset_hide_window)
        self.change_detected_timer.setSingleShot(True)
        
        self.auto_hide_timer = QTimer(self)
        self.auto_hide_timer.timeout.connect(self.auto_hide_to_tray)
        self.auto_hide_timer.setSingleShot(True)
        
        # 放大显示窗口 - 放在原监控画面上方
        self.magnify_window = MagnifyWindow(self)
        
        self.motion_active = False
        self.last_motion_time = time.time()
        self.magnify_update_interval = 0.2
        self.last_magnify_update = 0
        
        print(f"{datetime.now()} 监控程序启动 - ROI模式")
        print(f"ROI区域: x={self.roi['x']}, y={self.roi['y']}, w={self.roi['w']}, h={self.roi['h']}")
        print(f"显示分辨率: {self.display_width}x{self.display_height}")
        print(f"要隐藏的窗口: {self.window_name if self.window_name else '未设置'}")

    def load_settings(self):
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.window_name = settings.get("window_name", "")
        except Exception as e:
            print(f"加载设置失败: {e}")
            self.window_name = ""

    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("监控运行中")
        
        tray_menu = QMenu()
        
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)
        
        hide_action = QAction("隐藏到托盘", self)
        hide_action.triggered.connect(self.hide_window)
        tray_menu.addAction(hide_action)
        
        tray_menu.addSeparator()
        
        settings_action = QAction("设置隐藏窗口...", self)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)
        
        tray_menu.addSeparator()
        
        self.window_status_action = QAction("", self)
        self.window_status_action.setEnabled(False)
        tray_menu.addAction(self.window_status_action)
        self.update_window_status()
        
        tray_menu.addSeparator()
        
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def update_window_status(self):
        if self.window_name:
            hwnd = win32gui.FindWindow(None, self.window_name)
            if hwnd:
                self.window_status_action.setText(f"✓ 监控窗口: {self.window_name}")
            else:
                self.window_status_action.setText(f"⚠ 窗口未找到: {self.window_name}")
        else:
            self.window_status_action.setText("未设置隐藏窗口")

    def show_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_settings()
            self.update_window_status()
            print(f"窗口名称已更新为: {self.window_name}")

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()

    def show_window(self):
        self.show()
        self.setWindowState(Qt.WindowActive)
        self.raise_()
        self.activateWindow()
        
    def hide_window(self):
        self.hide()
        
    def auto_hide_to_tray(self):
        if not self.isHidden():
            self.hide_window()
            print(f"{datetime.now()} 5秒自动隐藏到托盘")

    def on_frame_received(self, frame):
        try:
            if frame is None:
                return
                
            self.frame_count += 1
            
            while not self.frame_queue.empty():
                try:
                    self.frame_queue.get_nowait()
                except queue.Empty:
                    break
            
            self.frame_queue.put(frame)
            
            self.fps_counter += 1
            now = time.time()
            if now - self.fps_last_time >= 1.0:
                self.fps = self.fps_counter
                self.fps_counter = 0
                self.fps_last_time = now
                
        except Exception as e:
            print(f"接收帧错误: {e}")

    def extract_roi(self, frame):
        try:
            x = max(0, min(self.roi["x"], frame.shape[1] - 1))
            y = max(0, min(self.roi["y"], frame.shape[0] - 1))
            w = min(self.roi["w"], frame.shape[1] - x)
            h = min(self.roi["h"], frame.shape[0] - y)
            
            if w <= 0 or h <= 0:
                return frame
            
            roi_frame = frame[y:y+h, x:x+w]
            return roi_frame
        except Exception as e:
            print(f"提取ROI错误: {e}")
            return frame

    def get_largest_motion_region(self, contours, frame_shape, roi_offset):
        if not contours:
            return None
        
        largest_contour = max(contours, key=cv2.contourArea)
        
        x_display, y_display, w_display, h_display = cv2.boundingRect(largest_contour)
        
        scale_x = self.roi_width / self.display_width
        scale_y = self.roi_height / self.display_height
        
        x_roi = int(x_display * scale_x)
        y_roi = int(y_display * scale_y)
        w_roi = int(w_display * scale_x)
        h_roi = int(h_display * scale_y)
        
        margin = 80
        x_roi = max(0, x_roi - margin)
        y_roi = max(0, y_roi - margin)
        w_roi = min(self.roi_width - x_roi, w_roi + margin * 2)
        h_roi = min(self.roi_height - y_roi, h_roi + margin * 2)
        
        x_original = self.roi["x"] + x_roi
        y_original = self.roi["y"] + y_roi
        
        return (x_original, y_original, w_roi, h_roi)

    def update_display(self):
        try:
            try:
                self.last_frame = self.frame_queue.get_nowait()
            except queue.Empty:
                if self.last_frame is None:
                    return
            
            if self.last_frame is None:
                return
            
            frame = self.last_frame.copy()
            current_time = time.time()
            
            roi_frame = self.extract_roi(frame)
            
            if roi_frame.size == 0:
                return
            
            display_frame = cv2.resize(roi_frame, (self.display_width, self.display_height))
            
            detect_scale = 0.5
            detect_width = int(self.display_width * detect_scale)
            detect_height = int(self.display_height * detect_scale)
            detect_frame = cv2.resize(display_frame, (detect_width, detect_height))
            detect_frame = cv2.GaussianBlur(detect_frame, (3, 3), 0)
            
            if self.initializing:
                self.bg_subtractor.apply(detect_frame)
                self.init_frames += 1
                if self.init_frames >= self.init_max_frames:
                    self.initializing = False
                    print(f"{datetime.now()} 背景初始化完成")
                
                cv2.putText(display_frame, f"Initializing... {self.init_frames}/{self.init_max_frames}", 
                           (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
                main_pixels = 0
                change_detected = False
            else:
                fgmask = self.bg_subtractor.apply(detect_frame)
                fgmask = cv2.medianBlur(fgmask, 3)
                fgmask = cv2.erode(fgmask, None, iterations=1)
                fgmask = cv2.dilate(fgmask, None, iterations=1)
                
                main_pixels = np.sum(fgmask == 255)
                change_detected = main_pixels > 2500
                
                if change_detected:
                    if not self.motion_active:
                        self.motion_active = True
                        print(f"{datetime.now()} 运动开始")
                    
                    self.last_motion_time = current_time
                    self.change_counter += 1
                else:
                    self.change_counter = 0
                    
                    if self.motion_active and (current_time - self.last_motion_time) > 1.0:
                        self.motion_active = False
                        print(f"{datetime.now()} 运动停止")
                
                if change_detected:
                    fgmask_big = cv2.resize(fgmask, (self.display_width, self.display_height))
                    
                    contours, _ = cv2.findContours(fgmask_big.astype(np.uint8), 
                                                   cv2.RETR_EXTERNAL, 
                                                   cv2.CHAIN_APPROX_SIMPLE)
                    
                    motion_contours = []
                    for contour in contours:
                        if cv2.contourArea(contour) > 500:
                            x, y, w, h = cv2.boundingRect(contour)
                            cv2.rectangle(display_frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
                            cv2.putText(display_frame, "Motion", (x, y-5),
                                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
                            motion_contours.append(contour)
                    
                    if motion_contours and (current_time - self.last_magnify_update) > self.magnify_update_interval:
                        roi_rect = self.get_largest_motion_region(motion_contours, frame.shape, 
                                                                  (self.roi["x"], self.roi["y"]))
                        if roi_rect:
                            self.magnify_window.update_magnified_view(frame, roi_rect)
                            self.last_magnify_update = current_time
                
                chg_text = f"Motion: {main_pixels//100}K pixels"
                cv2.putText(display_frame, chg_text, (10, 30), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
                
                status_text = "MOVEMENT" if change_detected else "IDLE"
                status_color = (0, 0, 255) if change_detected else (0, 255, 0)
                cv2.putText(display_frame, status_text, (10, 60), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.6, status_color, 2)
                
                if change_detected and self.change_counter == 1:
                    estimated_pixels = int(main_pixels / (detect_scale * detect_scale))
                    self.on_change_detected(estimated_pixels)
            
            fps_text = f"FPS: {self.fps}"
            cv2.putText(display_frame, fps_text, (10, self.display_height-20), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
            
            time_text = datetime.now().strftime("%H:%M:%S")
            time_size = cv2.getTextSize(time_text, cv2.FONT_HERSHEY_SIMPLEX, 0.7, 2)[0]
            cv2.putText(display_frame, time_text, 
                       (self.display_width - time_size[0] - 10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            
            if self.hide_windows:
                cv2.putText(display_frame, "HIDDEN", 
                           (self.display_width - 100, 60), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
            
            cv2.rectangle(display_frame, (0, 0), (self.display_width-1, self.display_height-1), 
                         (100, 100, 100), 1)
            
            rgb = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGB)
            h, w, ch = rgb.shape
            bytes_per_line = ch * w
            qimg = QImage(rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.label.setPixmap(pixmap)
            
        except Exception as e:
            print(f"显示更新错误: {e}")

    def on_change_detected(self, score):
        print(f"{datetime.now()} 检测到变化！score={score}")
        
        if self.isHidden():
            self.show_window()
            self.flash_window()
        
        self.auto_hide_timer.stop()
        self.auto_hide_timer.start(5000)
        
        if not self.hide_windows:
            self.hide_windows = True
            self.hide_other_window()
            winsound.Beep(1000, 150)

        self.change_detected_timer.start(8000)
        self.setWindowTitle("监控-HIDDEN")

    def flash_window(self):
        def toggle_red():
            if hasattr(self, '_flash_count') and self._flash_count > 0:
                self.setStyleSheet("background-color: red;" if self._flash_count % 2 == 1 else "")
                self._flash_count -= 1
                QTimer.singleShot(180, toggle_red)
            else:
                self.setStyleSheet("")

        self._flash_count = 8
        toggle_red()

    def hide_other_window(self):
        if not self.window_name:
            return
            
        try:
            self.hwnd = win32gui.FindWindow(None, self.window_name)
            
            if not self.hwnd:
                def enum_windows_callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if title and self.window_name.lower() in title.lower():
                            windows.append(hwnd)
                    return True
                
                windows = []
                win32gui.EnumWindows(enum_windows_callback, windows)
                if windows:
                    self.hwnd = windows[0]
            
            if self.hwnd:
                win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
                window_text = win32gui.GetWindowText(self.hwnd)
                print(f"{datetime.now()} 隐藏窗口: {window_text}")
            else:
                print(f"{datetime.now()} 未找到窗口: {self.window_name}")
                
        except Exception as e:
            print(f"隐藏窗口失败: {e}")

    def reset_hide_window(self):
        self.hide_windows = False
        if self.hwnd and self.window_name:
            try:
                if win32gui.IsWindow(self.hwnd):
                    win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
                    window_text = win32gui.GetWindowText(self.hwnd)
                    print(f"{datetime.now()} 恢复窗口: {window_text}")
            except Exception as e:
                print(f"恢复窗口失败: {e}")
        self.setWindowTitle("监控 - ROI模式")

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.quit_app()
        super().keyPressEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def changeEvent(self, event):
        if event.type() == event.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                QTimer.singleShot(100, self.hide_window)
        super().changeEvent(event)

    def closeEvent(self, event):
        if self.tray_icon and self.tray_icon.isVisible():
            self.hide_window()
            event.ignore()
        else:
            self.quit_app()

    def quit_app(self):
        print(f"{datetime.now()} 正在退出程序...")
        if hasattr(self, 'magnify_window') and self.magnify_window:
            self.magnify_window.close()
        if hasattr(self, 'capture_thread') and self.capture_thread:
            self.capture_thread.stop()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    player = VideoPlayer()
    sys.exit(app.exec_())
