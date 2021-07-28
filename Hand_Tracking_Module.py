import cv2
import mediapipe as mp
import time


class Hand_Detector():

    # Constructor for the defining the variables
    def __init__(self, mode=False, maxHands=1, detectionCon=0.5, trackCon=0.5):
        self.mode = mode
        self.maxHands = maxHands        # Here according to the mediapipe library we can detect only 2 hands
                                        # in one frame
        self.detectionCon = detectionCon        # This is Detection Confidence or Detection Accuracy
        self.trackCon = trackCon                # This is Tracking Confidence or Tracking Accuracy
        self.mpHands = mp.solutions.hands
        self.hands = self.mpHands.Hands(self.mode, self.maxHands, self.detectionCon, self.trackCon)
        self.mpDraw = mp.solutions.drawing_utils

    # Method for defining the points on the hand
    def findHands(self, frame, draw=True):
        imgRGB = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)     # Convert BGR to RGB
        self.Result = self.hands.process(imgRGB)
        if self.Result.multi_hand_landmarks:                # To detect Landmarks on the hand
            for handlms in self.Result.multi_hand_landmarks:
                # To put the landmarks on the hand and connect them
                if draw:
                    self.mpDraw.draw_landmarks(frame, handlms, self.mpHands.HAND_CONNECTIONS)
        # Return the final image with the landmarks
        return frame

    # Method to find the coordinates of the hand and id no. of the landmarks
    def find_Position(self, frame, handNo=0, draw=True):
        lmList = []         # Create a list of the landmarks and coordinates
        if self.Result.multi_hand_landmarks:
            myHand = self.Result.multi_hand_landmarks[handNo]
            for id, lm in enumerate(myHand.landmark):
                h, w, c = frame.shape           # This will give the height, width and Channel of the hand
                cx, cy = int(lm.x * w), int(lm.y * h)       # Convert the coordinates from decimals to pixels
                lmList.append([id, cx, cy])     # Add it to the List

                # Mark the necessary landmarks
                if draw:
                    cv2.circle(frame, (cx, cy), 8, (255, 0, 0), cv2.FILLED)
        # Return the list containing the landmarks and their coordinates
        return lmList


# Dummy Function from which we can copy the contents onto another project

def main():
    previous_time = 0           # for FPS
    current_time = 0            # for FPS
    cap = cv2.VideoCapture(0)   # Open the Video Cam
    detector = Hand_Detector()  # detector is the object of the class Hand_Detector
    while True:
        chk, frame = cap.read()
        frame = detector.findHands(frame)  # Calling the findHands method and storing the
                                           # output of the method in the "frame" variable
        List = detector.find_Position(frame)  # Calling the find_Position method and storing the position
                                              # Returned as list in "List" Variable
        if len(List) != 0:
             print(List[4])                  # Print only the x and y coordinates of id no. 4

        # To Print the FPS on the Screen; This FPS process can be skipped if needed; This is just for show purpose
        current_time = time.time()
        fps = 1 / (current_time - previous_time)
        previous_time = current_time
        cv2.putText(frame, str(int(fps)), (10, 30), cv2.FONT_HERSHEY_PLAIN, 2, (0, 0, 255), 3)

        # To open show the live Webcam feed
        cv2.imshow("VIDEO", frame)
        cv2.waitKey(1)


# To call the dummy function
if __name__ == "__main__":
    main()
