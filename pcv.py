import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttk 
import cv2
from PIL import Image, ImageTk
import mediapipe as mp
import numpy as np
import time
import threading
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

class WebcamApp:
    def __init__(self, root):
        
        # GUI
        self.root = root
        self.root.title(" Volume Controller ")
        self.root.resizable(False, False)
        style = ttk.Style()
        style.theme_use("cosmo")  #https://ttkbootstrap.readthedocs.io/en/latest/themes/light/

        # Webcam 
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            raise RuntimeError("Could not open the camera. Please check if it's connected.")

        # Create main frame
        self.frame = ttk.Frame(root)
        self.frame.pack()

        # Canvas for displaying frames
        self.canvas = tk.Canvas(self.frame, width=int(self.cap.get(4)) - 100, height=int(self.cap.get(4)) - 100)
        self.canvas.grid(row=0, column=0)


        # Meter for displaying volume
        volume_frame = ttk.Frame(self.frame)
        volume_frame.grid(row=0, column=1)
    
        self.meter = ttk.Meter(
            # https://ttkbootstrap.readthedocs.io/en/latest/api/widgets/meter/#ttkbootstrap.widgets.Meter
            volume_frame,
            metersize=180,
            padding=10,
            metertype="semi",
            textfont='-size 25 -weight bold',
            subtext="Volume Level",
            amounttotal=100,
            stripethickness=10,
            bootstyle='success'
        )
        self.meter.pack(fill=tk.BOTH, expand=tk.YES, padx=10, pady=10)

        # Label 
        self.status_label = tk.Label(self.frame, text="Stop", font="-size 10 -weight bold", bg="black", fg="red")
        self.status_label.place(x=476, y=275)

        # Audio Control 
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        self.volMin, self.volMax = self.volume.GetVolumeRange()[:2]

        # Hand Tracking 
        mpHands = mp.solutions.hands
        self.hands = mpHands.Hands()

        # Parameters for Hand Tracking
        self.alpha, self.smooth_vol = 0.8, 0
        self.is_running, self.is_paused = False, False
        self.start_time = time.time()
        self.buffer_size, self.max_change_percent = 5, 10
        self.length_multiplier = 2
        self.beta = 0.5
        self.volume_buffer = [0] * self.buffer_size


        # Start capturing frames
        self.img = None
        self.update()
        
        # Volume meter setup
        self.volume_thread = self.VolumeThread(self.get_system_volume, self.on_volume_change)
        self.volume_thread.start()

    def get_system_volume(self):
        volume_level = self.volume.GetMasterVolumeLevelScalar() * 100
        return volume_level

    def update(self):
        
        try:
            success, img = self.cap.read()
            img = cv2.flip(img, 1)
            img = cv2.resize(img, (640, 480))
            imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

            results = self.hands.process(imgRGB)

            if results.multi_hand_landmarks:
                handLandmark = results.multi_hand_landmarks[0]
                lmList = [[id, lm.x * img.shape[1], lm.y * img.shape[0]] for id, lm in enumerate(handLandmark.landmark)]
                finger_count = sum(1 for i, j in zip([4, 8, 12, 16, 20], [3, 7, 11, 15, 19]) if lmList[i][2] < lmList[j][2])

                if finger_count == 5 and not self.is_running:
                    self.start_time, self.is_running = time.time(), True
                elif finger_count == 3:
                    self.is_running, self.is_paused = False, True

                if self.is_running and time.time() - self.start_time > 1.2:
                    x1, y1, x2, y2 = lmList[4][1], lmList[4][2], lmList[8][1], lmList[8][2]
                    cv2.circle(img, (int(x1), int(y1)), 10, (0, 255, 0), cv2.FILLED)
                    cv2.circle(img, (int(x2), int(y2)), 10, (0, 255, 0), cv2.FILLED)
                    cv2.line(img, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), 2)

                    center_x1, center_y1, center_x2, center_y2 = lmList[4][1], lmList[4][2], lmList[8][1], lmList[8][2]
                    length = np.hypot(center_x2 - center_x1, center_y2 - center_y1)
                    target_vol = np.interp(length * self.length_multiplier, [30, 150], [self.volMin, self.volMax])

                    self.volume_buffer.pop(0)
                    self.volume_buffer.append(target_vol)

                    target_vol = int(np.average(self.volume_buffer, weights=np.arange(1, self.buffer_size + 1)) * self.beta + self.smooth_vol * (1 - self.beta))

                    self.smooth_vol = self.alpha * self.smooth_vol + (1 - self.alpha) * target_vol
                    self.volume.SetMasterVolumeLevel(self.smooth_vol, None)

            imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            self.img = ImageTk.PhotoImage(image=Image.fromarray(imgRGB))
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img)
            self.root.iconbitmap('icon.ico')

            self.root.after(10, self.update)
        except Exception as e:
            print(f"An error occurred: {e}")
            self.quit()

    class VolumeThread(threading.Thread):
        def __init__(self, get_volume_callback, on_change_callback):
            super().__init__()
            self.get_volume_callback = get_volume_callback
            self.on_change_callback = on_change_callback
            self.running = True

        def run(self):
            while self.running:
                volume = self.get_volume_callback()
                self.on_change_callback(volume)

    def on_volume_change(self, volume):
        self.rounded_volume = round(volume)

        if self.rounded_volume < 31:
            self.meter.configure(amountused=self.rounded_volume, bootstyle='danger')
            self.update_status_text()
        elif 30 <= self.rounded_volume <= 71:
            self.meter.configure(amountused=self.rounded_volume, bootstyle='warning')
            self.update_status_text()
        else:
            self.meter.configure(amountused=self.rounded_volume, bootstyle='success')
            self.update_status_text()

    def update_status_text(self):
        if self.is_running:
            self.status_label.config(text="Running", fg="green")
        else:
            self.status_label.config(text="Stop", fg="red")
        


root = ttk.Window()
app = WebcamApp(root)
root.mainloop()
app.volume_thread.running = False
app.volume_thread.join()
