import sys
import cv2
import win32gui
import win32con
import numpy as np
import winsound
from datetime import datetime
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import QTimer, Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QSystemTrayIcon, QStyle
import time
import queue
import os


class VideoCaptureThread(QThread):
    frame_ready = pyqtSignal(object)
    
    def __init__(self, camera_url):
        super().__init__()
        self.camera_url = camera_url
        self.running = True
        
    def run(self):
        os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = "rtsp_transport;tcp|buffer_size;1024|max_delay;0"
        
        cap = cv2.VideoCapture(self.camera_url, cv2.CAP_FFMPEG)
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        cap.set(cv2.CAP_PROP_FPS, 30)
        
        while self.running:
            try:
                ret, frame = cap.read()
                if ret:
                    self.frame_ready.emit(frame)
                else:
                    print(f"{datetime.now()} 读取失败，重新连接...")
                    cap.release()
                    time.sleep(1)
                    cap = cv2.VideoCapture(self.camera_url, cv2.CAP_FFMPEG)
                    cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
            except Exception as e:
                print(f"捕获异常: {e}")
                time.sleep(1)
        
        cap.release()
    
    def stop(self):
        self.running = False
        self.wait()


class VideoPlayer(QWidget):
    def __init__(self):
        super().__init__()

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
        
        # 创建视频捕获线程
        self.capture_thread = VideoCaptureThread(self.camera_url)
        self.capture_thread.frame_ready.connect(self.on_frame_received)
        self.capture_thread.start()

        self.hide_windows = False
        self.hwnd = -1

        # UI设置 - 窗口大小等于ROI显示大小
        self.resize(self.display_width, self.display_height)
        self.move(1700, 850)
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

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QApplication.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("监控运行中")
        self.tray_icon.show()

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
        
        print(f"{datetime.now()} 监控程序启动 - ROI模式")
        print(f"ROI区域: x={self.roi['x']}, y={self.roi['y']}, w={self.roi['w']}, h={self.roi['h']}")
        print(f"显示分辨率: {self.display_width}x{self.display_height}")

    def on_frame_received(self, frame):
        """收到新帧"""
        try:
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
        # 确保ROI在图像范围内
        x = max(0, self.roi["x"])
        y = max(0, self.roi["y"])
        w = min(self.roi["w"], frame.shape[1] - x)
        h = min(self.roi["h"], frame.shape[0] - y)
        
        # 提取ROI
        roi_frame = frame[y:y+h, x:x+w]
        
        return roi_frame

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
            qimg = QImage(rgb.data, w, h, ch * w, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.label.setPixmap(pixmap)
            
        except Exception as e:
            print(f"显示更新错误: {e}")

    def on_change_detected(self, score):
        print(f"{datetime.now()} 检测到变化！score={score}")
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

            self.flash_window()

        self.change_detected_timer.start(8000)
        self.setWindowTitle("监控-HIDDEN")

    def flash_window(self):
        def toggle_red():
            if hasattr(self, '_flash_count') and self._flash_count > 0:
                self.setStyleSheet("background-color: red;" if self._flash_count % 2 == 1 else "background-color: black;")
                self._flash_count -= 1
                QTimer.singleShot(180, toggle_red)
            else:
                self.setStyleSheet("")

        self._flash_count = 8
        toggle_red()

    def hide_other_window(self):
        window_title = "无标题 - 记事本"
        self.hwnd = win32gui.FindWindow(None, window_title)
        if self.hwnd:
            win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)

    def reset_hide_window(self):
        self.hide_windows = False
        if self.hwnd:
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
        self.setWindowTitle("监控 - ROI模式")

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        super().keyPressEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def closeEvent(self, event):
        self.capture_thread.stop()
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    player = VideoPlayer()
    sys.exit(app.exec_())