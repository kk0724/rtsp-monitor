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


class VideoCaptureThread(QThread):
    # 修复信号定义 - 使用正确的类型
    frame_ready = pyqtSignal(object)
    
    def __init__(self, camera_url):
        super().__init__()
        self.camera_url = camera_url
        self.running = True
        self.cap = None
        
    def run(self):
        # 设置OpenCV环境变量
        os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = "rtsp_transport;tcp;buffer_size=1024000"
        
        # 连接摄像头
        self.cap = cv2.VideoCapture(self.camera_url, cv2.CAP_FFMPEG)
        self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        self.cap.set(cv2.CAP_PROP_FPS, 30)
        
        # 等待摄像头连接
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
                    # 发送帧
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
        
        # 清理资源
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
        
        # 加载当前设置
        self.settings = self.load_settings()
        
        # 创建UI
        layout = QVBoxLayout()
        
        # 说明文字
        info_label = QLabel("设置要自动隐藏的窗口名称（支持模糊匹配）：")
        layout.addWidget(info_label)
        
        # 窗口名称输入
        form_layout = QFormLayout()
        self.window_name_edit = QLineEdit()
        self.window_name_edit.setText(self.settings.get("window_name", ""))
        self.window_name_edit.setPlaceholderText("例如：无标题 - 记事本、微信、Chrome")
        form_layout.addRow("窗口名称:", self.window_name_edit)
        layout.addLayout(form_layout)
        
        # 提示标签
        hint_label = QLabel("提示：可以输入窗口标题的一部分，程序会自动查找匹配的窗口")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        hint_label.setWordWrap(True)
        layout.addWidget(hint_label)
        
        # 当前窗口列表
        self.window_list_btn = QPushButton("查看当前打开的窗口")
        self.window_list_btn.clicked.connect(self.show_window_list)
        layout.addWidget(self.window_list_btn)
        
        # 测试按钮
        test_btn = QPushButton("测试查找窗口")
        test_btn.clicked.connect(self.test_find_window)
        layout.addWidget(test_btn)
        
        # 按钮区域
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
        """加载设置"""
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载设置失败: {e}")
        return {"window_name": ""}
    
    def save_settings(self):
        """保存设置"""
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
        """测试查找窗口"""
        window_name = self.window_name_edit.text().strip()
        if not window_name:
            QMessageBox.warning(self, "提示", "请输入窗口名称")
            return
        
        # 查找窗口 - 精确匹配
        hwnd = win32gui.FindWindow(None, window_name)
        if hwnd:
            # 获取窗口标题
            window_text = win32gui.GetWindowText(hwnd)
            QMessageBox.information(self, "成功", f"找到精确匹配窗口：{window_text}")
            return
        
        # 模糊匹配
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):  # 只显示可见窗口
                title = win32gui.GetWindowText(hwnd)
                if title and window_name.lower() in title.lower():
                    windows.append(f"{title} (句柄: {hwnd})")
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            msg = "找到以下匹配窗口：\n" + "\n".join(windows[:10])  # 最多显示10个
            QMessageBox.information(self, "成功", msg)
        else:
            QMessageBox.warning(self, "失败", "未找到匹配的窗口")
    
    def show_window_list(self):
        """显示当前打开的窗口列表"""
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # 只显示有标题的窗口
                    windows.append(f"{title}")
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            msg = "当前打开的窗口：\n" + "\n".join(sorted(windows)[:20])  # 最多显示20个
            QMessageBox.information(self, "窗口列表", msg)
        else:
            QMessageBox.information(self, "窗口列表", "没有找到可见窗口")


class VideoPlayer(QWidget):
    def __init__(self):
        super().__init__()

        # 忽略libpng警告
        os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.png.warning=false'
        
        self.camera_url = 'rtsp://admin:guest123@192.168.6.105:554/Streaming/Channels/101'
        
        # ROI区域 (用户圈出来的范围)
        self.roi = {
            "x": 360,
            "y": 400,
            "w": 1170,
            "h": 1040
        }
        
        # ROI的宽高
        self.roi_width = self.roi["w"]
        self.roi_height = self.roi["h"]
        
        # 显示窗口大小 (保持ROI的宽高比，但缩小显示)
        self.display_width = 640
        self.display_height = int(self.display_width * self.roi_height / self.roi_width)
        
        # FPS统计
        self.fps = 0
        self.fps_counter = 0
        self.fps_last_time = time.time()
        
        # 帧队列
        self.frame_queue = queue.Queue(maxsize=2)
        self.last_frame = None
        self.frame_count = 0
        
        print(f"{datetime.now()} 初始化视频播放器...")
        print(f"摄像头地址: {self.camera_url}")
        
        # 创建视频捕获线程
        self.capture_thread = VideoCaptureThread(self.camera_url)
        self.capture_thread.frame_ready.connect(self.on_frame_received)
        self.capture_thread.start()

        self.hide_windows = False
        self.hwnd = -1
        self.window_name = ""  # 要隐藏的窗口名称
        
        # 加载设置
        self.load_settings()

        # UI设置 - 窗口大小等于ROI显示大小
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

        # 创建系统托盘图标
        self.setup_tray_icon()

        # 运动检测
        self.bg_subtractor = cv2.createBackgroundSubtractorMOG2(
            history=100,
            varThreshold=36,
            detectShadows=False
        )

        self.change_counter = 0
        self.min_consecutive_changes = 2

        self.initializing = True
        self.init_frames = 0
        self.init_max_frames = 50

        # 显示定时器
        self.display_timer = QTimer(self)
        self.display_timer.timeout.connect(self.update_display)
        self.display_timer.start(20)

        self.show()
        
        self.change_detected_timer = QTimer(self)
        self.change_detected_timer.timeout.connect(self.reset_hide_window)
        self.change_detected_timer.setSingleShot(True)
        
        # 新增：自动隐藏定时器
        self.auto_hide_timer = QTimer(self)
        self.auto_hide_timer.timeout.connect(self.auto_hide_to_tray)
        self.auto_hide_timer.setSingleShot(True)
        
        print(f"{datetime.now()} 监控程序启动 - ROI模式")
        print(f"ROI区域: x={self.roi['x']}, y={self.roi['y']}, w={self.roi['w']}, h={self.roi['h']}")
        print(f"显示分辨率: {self.display_width}x{self.display_height}")
        print(f"要隐藏的窗口: {self.window_name if self.window_name else '未设置'}")

    def load_settings(self):
        """加载设置"""
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.window_name = settings.get("window_name", "")
        except Exception as e:
            print(f"加载设置失败: {e}")
            self.window_name = ""

    def setup_tray_icon(self):
        """设置系统托盘图标"""
        self.tray_icon = QSystemTrayIcon(self)
        # 使用内置图标避免libpng警告
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("监控运行中")
        
        # 创建托盘菜单
        tray_menu = QMenu()
        
        # 显示窗口动作
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)
        
        # 隐藏窗口动作（最小化到托盘）
        hide_action = QAction("隐藏到托盘", self)
        hide_action.triggered.connect(self.hide_window)
        tray_menu.addAction(hide_action)
        
        tray_menu.addSeparator()
        
        # 设置动作
        settings_action = QAction("设置隐藏窗口...", self)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)
        
        tray_menu.addSeparator()
        
        # 显示当前隐藏窗口状态
        self.window_status_action = QAction("", self)
        self.window_status_action.setEnabled(False)
        tray_menu.addAction(self.window_status_action)
        self.update_window_status()
        
        tray_menu.addSeparator()
        
        # 退出动作
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        
        # 托盘图标点击事件
        self.tray_icon.activated.connect(self.tray_icon_activated)
        
        self.tray_icon.show()

    def update_window_status(self):
        """更新窗口状态显示"""
        if self.window_name:
            # 检查窗口是否存在
            hwnd = win32gui.FindWindow(None, self.window_name)
            if hwnd:
                self.window_status_action.setText(f"✓ 监控窗口: {self.window_name}")
            else:
                self.window_status_action.setText(f"⚠ 窗口未找到: {self.window_name}")
        else:
            self.window_status_action.setText("未设置隐藏窗口")

    def show_settings(self):
        """显示设置对话框"""
        dialog = SettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            # 重新加载设置
            self.load_settings()
            self.update_window_status()
            print(f"窗口名称已更新为: {self.window_name}")

    def tray_icon_activated(self, reason):
        """托盘图标被点击时的处理"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()
        elif reason == QSystemTrayIcon.Trigger:  # 单击
            pass

    def show_window(self):
        """显示窗口"""
        self.show()
        self.setWindowState(Qt.WindowActive)
        self.raise_()
        self.activateWindow()
        
    def hide_window(self):
        """隐藏窗口到托盘（不显示提示）"""
        self.hide()
        
    def auto_hide_to_tray(self):
        """10秒后自动隐藏到托盘"""
        if not self.isHidden():  # 如果窗口还没被隐藏
            self.hide_window()
            print(f"{datetime.now()} 10秒自动隐藏到托盘")

    def on_frame_received(self, frame):
        """收到新帧"""
        try:
            if frame is None:
                return
                
            self.frame_count += 1
            
            # 清空队列，只保留最新帧
            while not self.frame_queue.empty():
                try:
                    self.frame_queue.get_nowait()
                except queue.Empty:
                    break
            
            self.frame_queue.put(frame)
            
            # 统计FPS
            self.fps_counter += 1
            now = time.time()
            if now - self.fps_last_time >= 1.0:
                self.fps = self.fps_counter
                self.fps_counter = 0
                self.fps_last_time = now
                
        except Exception as e:
            print(f"接收帧错误: {e}")

    def extract_roi(self, frame):
        """从原图中提取ROI区域"""
        try:
            # 确保ROI在图像范围内
            x = max(0, min(self.roi["x"], frame.shape[1] - 1))
            y = max(0, min(self.roi["y"], frame.shape[0] - 1))
            w = min(self.roi["w"], frame.shape[1] - x)
            h = min(self.roi["h"], frame.shape[0] - y)
            
            if w <= 0 or h <= 0:
                return frame
            
            # 提取ROI
            roi_frame = frame[y:y+h, x:x+w]
            
            return roi_frame
        except Exception as e:
            print(f"提取ROI错误: {e}")
            return frame

    def update_display(self):
        """更新显示 - 只显示ROI区域"""
        try:
            # 获取最新帧
            try:
                self.last_frame = self.frame_queue.get_nowait()
            except queue.Empty:
                if self.last_frame is None:
                    return
            
            if self.last_frame is None:
                return
            
            frame = self.last_frame.copy()
            
            # 提取ROI区域
            roi_frame = self.extract_roi(frame)
            
            if roi_frame.size == 0:
                return
            
            # 缩小ROI以适应显示窗口
            display_frame = cv2.resize(roi_frame, (self.display_width, self.display_height))
            
            # 进一步缩小用于运动检测（提高速度）
            detect_scale = 0.5
            detect_width = int(self.display_width * detect_scale)
            detect_height = int(self.display_height * detect_scale)
            detect_frame = cv2.resize(display_frame, (detect_width, detect_height))
            detect_frame = cv2.GaussianBlur(detect_frame, (3, 3), 0)
            
            # 运动检测
            if self.initializing:
                self.bg_subtractor.apply(detect_frame)
                self.init_frames += 1
                if self.init_frames >= self.init_max_frames:
                    self.initializing = False
                    print(f"{datetime.now()} 背景初始化完成")
                
                # 显示初始化状态
                cv2.putText(display_frame, f"Initializing... {self.init_frames}/{self.init_max_frames}", 
                           (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
                main_pixels = 0
                change_detected = False
            else:
                # 运动检测
                fgmask = self.bg_subtractor.apply(detect_frame)
                fgmask = cv2.medianBlur(fgmask, 3)
                fgmask = cv2.erode(fgmask, None, iterations=1)
                fgmask = cv2.dilate(fgmask, None, iterations=1)
                
                main_pixels = np.sum(fgmask == 255)
                change_detected = main_pixels > 3000  # 调整后的阈值
                
                if change_detected:
                    self.change_counter += 1
                    if self.change_counter >= self.min_consecutive_changes:
                        # 估算原始像素数
                        estimated_pixels = int(main_pixels / (detect_scale * detect_scale))
                        self.on_change_detected(estimated_pixels)
                        self.change_counter = 0
                else:
                    self.change_counter = 0
                
                # 在显示画面上标记检测区域（可选）
                if change_detected:
                    # 将检测到的运动区域映射回显示画面
                    fgmask_big = cv2.resize(fgmask, (self.display_width, self.display_height))
                    # 用红色半透明覆盖运动区域
                    display_frame[fgmask_big > 0] = display_frame[fgmask_big > 0] * 0.7 + np.array([0, 0, 255]) * 0.3
            
            # 显示检测信息
            if not self.initializing:
                chg_text = f"Motion: {main_pixels//100}K pixels"
                cv2.putText(display_frame, chg_text, (10, 30), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
                
                # 显示状态
                status_text = "MOVEMENT" if change_detected else "IDLE"
                status_color = (0, 0, 255) if change_detected else (0, 255, 0)
                cv2.putText(display_frame, status_text, (10, 60), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.6, status_color, 2)
            
            # 显示FPS
            fps_text = f"FPS: {self.fps}"
            cv2.putText(display_frame, fps_text, (10, self.display_height-20), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
            
            # 显示时间
            time_text = datetime.now().strftime("%H:%M:%S")
            time_size = cv2.getTextSize(time_text, cv2.FONT_HERSHEY_SIMPLEX, 0.7, 2)[0]
            cv2.putText(display_frame, time_text, 
                       (self.display_width - time_size[0] - 10, 30), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            
            # 显示隐藏状态
            if self.hide_windows:
                cv2.putText(display_frame, "HIDDEN", 
                           (self.display_width - 100, 60), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
            
            # 在ROI边缘画一个细框（可选）
            cv2.rectangle(display_frame, (0, 0), (self.display_width-1, self.display_height-1), 
                         (100, 100, 100), 1)
            
            # 转换并显示
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
        
        # 如果窗口是隐藏的，自动显示出来
        if self.isHidden():
            self.show_window()
            # 闪烁窗口吸引注意力
            self.flash_window()
        
        # 启动10秒自动隐藏定时器（先停止之前的）
        self.auto_hide_timer.stop()
        self.auto_hide_timer.start(10000)  # 10秒后自动隐藏
        
        if not self.hide_windows:
            self.hide_windows = True
            self.hide_other_window()

            winsound.Beep(1000, 150)
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)

            self.tray_icon.showMessage(
                "有人来了！",
                f"检测到运动 ({score//1000}k像素)",
                QSystemTrayIcon.Warning,
                10000
            )

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
        """隐藏其他窗口"""
        if not self.window_name:  # 如果没有设置窗口名称，就不隐藏
            return
            
        try:
            # 先尝试精确匹配
            self.hwnd = win32gui.FindWindow(None, self.window_name)
            
            # 如果没找到，尝试模糊匹配（遍历所有窗口）
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
                    self.hwnd = windows[0]  # 使用第一个匹配的窗口
            
            if self.hwnd:
                win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
                window_text = win32gui.GetWindowText(self.hwnd)
                print(f"{datetime.now()} 隐藏窗口: {window_text}")
            else:
                print(f"{datetime.now()} 未找到窗口: {self.window_name}")
                
        except Exception as e:
            print(f"隐藏窗口失败: {e}")

    def reset_hide_window(self):
        """恢复隐藏的窗口"""
        self.hide_windows = False
        if self.hwnd and self.window_name:  # 确保有设置窗口名称
            try:
                if win32gui.IsWindow(self.hwnd):  # 检查窗口是否还存在
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
        """处理窗口状态变化"""
        if event.type() == event.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                # 当窗口最小化时，隐藏到系统托盘（无提示）
                QTimer.singleShot(100, self.hide_window)
        super().changeEvent(event)

    def closeEvent(self, event):
        """处理关闭事件"""
        # 如果只是点击关闭按钮，则隐藏到托盘而不是退出（无提示）
        if self.tray_icon and self.tray_icon.isVisible():
            self.hide_window()  # 直接隐藏，不显示提示
            event.ignore()  # 忽略关闭事件
        else:
            self.quit_app()

    def quit_app(self):
        """退出应用程序"""
        print(f"{datetime.now()} 正在退出程序...")
        if hasattr(self, 'capture_thread') and self.capture_thread:
            self.capture_thread.stop()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # 防止最后一个窗口关闭时退出程序
    player = VideoPlayer()
    sys.exit(app.exec_())
