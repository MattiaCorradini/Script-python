from ctypes import cast, POINTER
import cv2
import mediapipe as mp
import time
import math
import numpy as np
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import sounddevice as sd

class HandDetector:
    def __init__(self, mode=False, max_hands=2, detection_confidence=0.5, tracking_confidence=0.5):
        self.mode = mode
        self.max_hands = max_hands
        self.detection_confidence = detection_confidence
        self.tracking_confidence = tracking_confidence

        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(
            static_image_mode=False,
            max_num_hands=self.max_hands,
            min_detection_confidence=self.detection_confidence,
            min_tracking_confidence=self.tracking_confidence
        )

        self.mp_draw = mp.solutions.drawing_utils
        self.tip_ids = [4, 8, 12, 16, 20]

    def find_hands(self, img, draw=True):
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        self.results = self.hands.process(img_rgb)

        if self.results.multi_hand_landmarks:
            for hand_landmarks in self.results.multi_hand_landmarks:
                if draw:
                    self.mp_draw.draw_landmarks(img, hand_landmarks, self.mp_hands.HAND_CONNECTIONS)

        return img

    def find_positions(self, img, hand_num=0, draw=True):
        x_list = []
        y_list = []
        bbox = []
        lm_list = []

        if self.results.multi_hand_landmarks:
            hand_landmarks = self.results.multi_hand_landmarks[hand_num]
            for id, landmark in enumerate(hand_landmarks.landmark):
                image_height, image_width, _ = img.shape
                x, y = int(landmark.x * image_width), int(landmark.y * image_height)
                x_list.append(x)
                y_list.append(y)
                lm_list.append([id, x, y])

                if draw:
                    cv2.circle(img, (x, y), 5, (255, 0, 255), cv2.FILLED)

            x_min, x_max = min(x_list), max(x_list)
            y_min, y_max = min(y_list), max(y_list)
            bbox = x_min, y_min, x_max, y_max

            if draw:
                cv2.rectangle(img, (bbox[0] - 20, bbox[1] - 20), (bbox[2] + 20, bbox[3] + 20), (0, 255, 0), 2)

        return lm_list, bbox

    def fingers_up(self):
        fingers = []
        if self.results.multi_hand_landmarks:
            hand_landmarks = self.results.multi_hand_landmarks[0]
            for tip_id in self.tip_ids:
                if hand_landmarks.landmark[tip_id].y < hand_landmarks.landmark[tip_id - 2].y:
                    fingers.append(1)
                else:
                    fingers.append(0)

        return fingers

    def find_distance(self, p1, p2, img, draw=True):
        x1, y1 = self.lm_list[p1][1], self.lm_list[p1][2]
        x2, y2 = self.lm_list[p2][1], self.lm_list[p2][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        if draw:
            cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
            cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
            cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
            cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)

        length = math.hypot(x2 - x1, y2 - y1)
        return length, img


def main():
    cap = cv2.VideoCapture(0)
    p_time = 0
    detector = HandDetector()

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume_min, volume_max, _ = volume.GetVolumeRange()

    while True:
        success, img = cap.read()
        img = detector.find_hands(img)
        lm_list, bbox = detector.find_positions(img)

        if len(lm_list) != 0:
            fingers = detector.fingers_up()
            volume_value = np.interp(lm_list[4][1], [bbox[0], bbox[2]], [0, 1])
            volume_value = max(0, min(volume_value, 1))
            volume_level = volume_min + (volume_max - volume_min) * volume_value
            volume.SetMasterVolumeLevel(volume_level, None)

            if fingers[1] == 1 and fingers[2] == 0:
                current_volume = volume.GetMasterVolumeLevelScalar()
                volume.SetMasterVolumeLevelScalar(min(current_volume + 0.1, 1), None)
            elif fingers[1] == 0 and fingers[2] == 1:
                current_volume = volume.GetMasterVolumeLevelScalar()
                volume.SetMasterVolumeLevelScalar(max(current_volume - 0.1, 0), None)

        c_time = time.time()
        fps = 1 / (c_time - p_time)
        p_time = c_time

        cv2.putText(img, f'FPS: {int(fps)}', (10, 70), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 255), 3)

        cv2.imshow("Image", img)
        cv2.waitKey(1)


if __name__ == "__main__":
    main()