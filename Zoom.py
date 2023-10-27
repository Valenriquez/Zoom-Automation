
import os              
import pandas as pd    
import pyautogui
import time
from datetime import datetime
from pywinauto import Application

os.startfile(r"C:\Users\enriq\AppData\Roaming\Zoom\bin\Zoom.exe")
time.sleep(3)

join_btn = pyautogui.locateCenterOnScreen(r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 160519.png")
if join_btn:
    pyautogui.click(join_btn)
else:
    print("Join button not found")
meetingidbtn=pyautogui.locateCenterOnScreen(r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 160851.png")
pyautogui.moveTo(meetingidbtn)
pyautogui.write("4244200467")
time.sleep(1)

image_paths = [
    r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 161451.png",
    r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 161511.png",
]
for image_path in image_paths:
    buttons = list(pyautogui.locateAllOnScreen(image_path))
    for btn in buttons:
        pyautogui.moveTo(btn)
        pyautogui.click()
        time.sleep(0.5)

joinblue_btn=pyautogui.locateCenterOnScreen(r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 163551.png")
pyautogui.moveTo(joinblue_btn)
pyautogui.click()
time.sleep(1)


passcode=pyautogui.locateCenterOnScreen(r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 162033.png")
pyautogui.moveTo(passcode)
pyautogui.write("1234567")

joinfinal_btn=pyautogui.locateCenterOnScreen(r"D:\\Mis archivos\\Descargas\\Captura de pantalla 2023-10-27 163730.png")
pyautogui.moveTo(joinfinal_btn)
pyautogui.click()
time.sleep(1)
1234567
