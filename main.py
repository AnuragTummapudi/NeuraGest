import cv2
import mediapipe as mp
import pyautogui
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(min_detection_confidence=0.7, min_tracking_confidence=0.7)
mp_draw = mp.solutions.drawing_utils
screen_width, screen_height = pyautogui.size()
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
click_threshold = 0.05
volume_threshold = 0.1
dragging = False
last_click_time = 0
scroll_speed = 5
previous_volume_distance = None
text_display = "NeuraGest by Anurag"

cap = cv2.VideoCapture(0)

while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break

    frame = cv2.flip(frame, 1)  # Mirror the frame
    frame_height, frame_width, _ = frame.shape

    # Add text to frame
    cv2.putText(frame, text_display, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

    # Convert to RGB for Mediapipe processing
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb_frame)

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            # Get landmark positions
            landmarks = hand_landmarks.landmark
            index_finger_tip = landmarks[8]
            thumb_tip = landmarks[4]
            middle_finger_tip = landmarks[12]
            ring_finger_tip = landmarks[16]
            pinky_tip = landmarks[20]

            # Convert to screen coordinates
            index_x = int(index_finger_tip.x * screen_width)
            index_y = int(index_finger_tip.y * screen_height)

            # Check if thumb is open (above index finger base)
            thumb_open = thumb_tip.x < landmarks[2].x  # Adjust this condition if needed

            # Check if all fingers are closed except the middle finger
            fingers_closed_except_middle = (
                index_finger_tip.y > landmarks[6].y and
                ring_finger_tip.y > landmarks[14].y and
                pinky_tip.y > landmarks[18].y and
                middle_finger_tip.y < landmarks[10].y
            )

            # Move cursor only if thumb is closed
            if not thumb_open:
                pyautogui.moveTo(index_x, index_y, duration=0.1)

            # If middle finger is open and all other fingers are closed, close all tabs except the frame tab
            if fingers_closed_except_middle:
                pyautogui.hotkey("ctrl", "shift", "w")  # Close all tabs except the main window
                text_display = "Same to You"
                time.sleep(1)  # Prevent continuous triggering
            else:
                text_display = "NeuraGest by Anurag"  # Reset text

            # Draw a circle at the index finger tip
            cv2.circle(frame, (int(index_finger_tip.x * frame_width), int(index_finger_tip.y * frame_height)), 10, (0, 255, 255), -1)

            # Distance between middle finger and thumb for left-click
            distance_left_click = abs(middle_finger_tip.x - thumb_tip.x) + abs(middle_finger_tip.y - thumb_tip.y)

            # Distance between ring finger and thumb for right-click
            distance_right_click = abs(ring_finger_tip.x - thumb_tip.x) + abs(ring_finger_tip.y - thumb_tip.y)

            # Distance between index and thumb for volume control
            distance_volume = abs(index_finger_tip.x - thumb_tip.x) + abs(index_finger_tip.y - thumb_tip.y)

            # Left Click Gesture: Middle finger and thumb touch
            if distance_left_click < click_threshold:
                if time.time() - last_click_time > 0.3:
                    pyautogui.click()
                    last_click_time = time.time()

            # Right Click Gesture: Ring finger and thumb touch
            if distance_right_click < click_threshold:
                if time.time() - last_click_time > 0.3:
                    pyautogui.rightClick()
                    last_click_time = time.time()

            # Drag Gesture: Hold middle finger + thumb to start dragging
            if distance_left_click < click_threshold and not dragging:
                pyautogui.mouseDown()
                dragging = True
            elif distance_left_click > click_threshold and dragging:
                pyautogui.mouseUp()
                dragging = False


            # Volume Control Gesture: Thumb + index finger spreading apart
            if previous_volume_distance is not None:
                volume_diff = distance_volume - previous_volume_distance
                if abs(volume_diff) > volume_threshold:
                    current_volume = volume.GetMasterVolumeLevelScalar()
                    new_volume = max(0.0, min(1.0, current_volume + (volume_diff * 2)))
                    volume.SetMasterVolumeLevelScalar(new_volume, None)

            previous_volume_distance = distance_volume

    cv2.imshow("Neuro Gest", frame)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
