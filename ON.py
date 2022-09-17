import time
import pyautogui

for i in range(1000):
    time.sleep(60)
    pyautogui.typewrite(str(i+1) + "min has passed")
    pyautogui.hotkey("enter")
