import cv2
import time                                     # for FPS (frames per second)
import numpy as np                              # for converting our hand motion to volume control
import Hand_Tracking_Module as htm              # Custom created module
import math                                     # for finding the hypotenuse between our fingers
import pycaw                                    # for Controlling the system Volume

# Usage of pycaw and importing necessary modules
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Opening the Webcam; Use this statement if using external camera: cap = cv2.VideoCapture(1)
cap = cv2.VideoCapture(0)
# Setting the Window size of the WebCam display
cap.set(3, 640)
cap.set(4, 480)

# For Fps
previous_time = 0

# Object of the Hand_Detector Class of the Hand_Tracking_Module
detector = htm.Hand_Detector(detectionCon=0.75)

# This is from the pycaw library
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
Volume_Range = volume.GetVolumeRange()

#  To get the minimum and Maximum system volume which is usually -65.25 and 0.0 respectively in Windows
minVol = Volume_Range[0]
maxVol = Volume_Range[1]
vol = 0

while True:
    # To read the Webcam feed as images
    check, frame = cap.read()

    # calling the method findHands
    frame = detector.findHands(frame)

    # Storing the position of the hand
    lmList = detector.find_Position(frame, draw=False)
    if len(lmList) != 0:
        x1, y1 = lmList[4][1], lmList[4][2]     # To find the tip of the Thumb
        x2, y2 = lmList[8][1], lmList[8][2]     # To find the tip of the Index finger
        cx, cy = (x1+x2)//2, (y1+y2)//2         # To find the center of the thumb and Index finger

        # To draw a circle around the Thumb and Index Finger
        cv2.circle(frame, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(frame, (x2, y2), 15, (255, 0, 255), cv2.FILLED)

        # To draw a line to connect the thumb and Index Fingers
        cv2.line(frame, (x1, y1), (x2, y2), (0, 0, 255), 2)

        # TO draw a circle around the middle of the thumb and Index finger
        cv2.circle(frame, (cx, cy), 7, (0, 0, 0), cv2.FILLED)

        # to find the distance between the thumb and index fingers
        length = math.hypot(x2-x1, y2-y1)

        # to match the volume ranges obtained from pycaw package and the distance ranges between our fingers
        vol = np.interp(length, [50, 150], [minVol, maxVol])

        # To change the colour of the mid point circle when the distance is minimum to RED
        if length < 50:
            cv2.circle(frame, (cx, cy), 7, (0, 0, 255), cv2.FILLED)
        # To change the colour of the mid point circle when the distance is maximum to Blue
        elif length > 150:
            cv2.circle(frame, (cx, cy), 15, (255, 0, 0), cv2.FILLED)

        # To Adjust the system Volume according to the distance between our index and thumb finger
        volume.SetMasterVolumeLevel(vol, None)

    # To display the fps on the screen
    current_time = time.time()
    fps = 1/(current_time - previous_time)
    previous_time = current_time
    cv2.putText(frame, f"FPS: {str(int(fps))}", (10, 50), cv2.FONT_HERSHEY_PLAIN, 2, (0, 0, 255), 3)

    # To display Custom message on the screen
    cv2.putText(frame, f"Please see your System Volume", (10, 70), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 255), 3)

    # To display the WebCam feed
    cv2.imshow("Volume_Hand_Control", frame)
    cv2.waitKey(1)


volume.SetMasterVolumeLevel(0.0, None)