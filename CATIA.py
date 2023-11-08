import cv2
import mediapipe as mp
import time
import win32com.client   # pip install pypiwin32
import math

catia = win32com.client.Dispatch('catia.application')  # Capture CATIA application API
doc = catia.activedocument.Product.GetTechnologicalObject('Mechanisms').Item(1)  # Read the hand object in CATIA

def angle(l1, l2, l3):  #  Calculate joint angles using cosine law
    cos = (l1 * l1 + l2 * l2 - l3 * l3) / 2 / l1 / l2
    cos_angle = 180 - math.acos(cos) * 180 / 3.1415926535
    if cos_angle >= 80:
        return 80  # Exeed the upper boundary
    else:
        return cos_angle  # Return real-time angle


def length(a, b):  # Calculate distance between 2 points
    x = x_list[a] - x_list[b]
    y = y_list[a] - y_list[b]
    z = z_list[a] - z_list[b]
    l = math.sqrt(x * x + y * y + z * z)
    return l


cap = cv2.VideoCapture(0)  # Turn on the camera
mpHands = mp.solutions.hands
hands = mpHands.Hands(static_image_mode=False,
                      max_num_hands=2,
                      min_detection_confidence=0.5,
                      min_tracking_confidence=0.5)  # Object instantiation
mpDraw = mp.solutions.drawing_utils
pTime = 0
cTime = 0
angles = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  
doc.putcommandvalues(angels)   # Initialize CATIA DMU

while True:
    x_list = []
    y_list = []
    z_list = []  # Clear x,y,z list
    success, img = cap.read()
    img = cv2.flip(img, 1)  
    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(imgRGB)
    landmarks_list = results.multi_hand_landmarks
    if results.multi_hand_landmarks:
        for hands2 in results.multi_hand_landmarks:
            for id, hand_lms in enumerate(hands2.landmark):  # Extract coordinates
                x_list.append(hand_lms.x)
                y_list.append(hand_lms.y)
                z_list.append(hand_lms.z)

        angles[0] = angle(length(1, 2), length(2, 3), length(1, 3))
        angles[1] = angle(length(2, 3), length(3, 4), length(2, 4))
        angles[2] = angle(length(0, 5), length(5, 6), length(0, 6))
        angles[3] = angle(length(6, 5), length(7, 6), length(5, 7))
        angles[4] = angle(length(6, 7), length(7, 8), length(6, 8))
        angles[5] = angle(length(0, 9), length(9, 10), length(0, 10))
        angles[6] = angle(length(10, 9), length(11, 10), length(9, 11))
        angles[7] = angle(length(10, 11), length(11, 12), length(10, 12))
        angles[8] = angle(length(0, 13), length(13, 14), length(0, 14))
        angles[9] = angle(length(14, 13), length(15, 14), length(13, 15))
        angles[10] = angle(length(14, 15), length(15, 16), length(14, 16))
        angles[11] = angle(length(0, 17), length(17, 18), length(0, 18))
        angles[12] = angle(length(18, 17), length(19, 18), length(17, 19))
        angles[13] = angle(length(18, 19), length(19, 20), length(18, 20))  # Calculate angles with three given coordinates
        doc.putcommandvalues(angles)  # Send the angle of each joint respectively and refresh.
