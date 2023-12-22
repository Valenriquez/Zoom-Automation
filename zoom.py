
import os              
import pandas as pd    
import pyautogui
import requests
from io import BytesIO
from PIL import Image
import time
from datetime import datetime
import sys


# Read the Excel file
file_path = file_path = "D:\\Mis archivos\\Descargas\\schedule.xlsx"
sheet_name = 'Schedule'
df = pd.read_excel(file_path, sheet_name=sheet_name)

for index, row in df.iterrows():
    scheduled_date = row['Date']
    scheduled_hour = row['Hour']
    meetingNumber = row['Meeting Number']
    Password = row['Password']

    # Parse 'Date' and 'Hour' columns
    date_str = scheduled_date.strftime('%Y-%m-%d')
    time_str = scheduled_hour.strftime('%H:%M:%S')
    meeting_number_str = str(meetingNumber)
    password_str = str(Password)

    # Combine date and time strings into a datetime object
    scheduled_datetime = datetime.strptime(f"{date_str} {time_str}", '%Y-%m-%d %H:%M:%S')
    print(scheduled_datetime)
    
    # Convert string to datetime object
    #formatted_datetime = scheduled_datetime.strftime('%d/%m/%Y %H:%M:%S')

    # Check if the scheduled datetime is in the future
    current_datetime = datetime.now()
    print(current_datetime)

    if scheduled_datetime > current_datetime:
        time_until_run = scheduled_datetime - current_datetime
        print(f"Scheduled for {scheduled_datetime}. Time until run: {time_until_run}")
        sys.exit(1) 

# Function to download images from URLs and return their local paths
def download_image(url):
    response = requests.get(url)
    if response.status_code == 200:
        image_content = response.content
        try:
            image = Image.open(BytesIO(image_content))
            local_path = f"local_image_{url.split('/')[-1].split('?')[0]}.png"  # Create a unique local path for each image
            image.save(local_path)
            return local_path
        except Exception as e:
            print(f"Error opening image: {e}")
    else:
        print(f"Failed to download image from {url}. Status code: {response.status_code}")

os.startfile(r"C:\Users\enriq\AppData\Roaming\Zoom\bin\Zoom.exe")
time.sleep(3)

image_info = {
    "PlusAddMeeting": "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/PlusAddMeeting.png",
    "IDMeeting" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/IDMeeting.png",
    "Deactivate Audio" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/DeactivateAudio.png" ,
    "Deactivate Video" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/DeactivateVideo.png" ,
    "Join" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/Join.png" ,
    "AccessCodeMeeting" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/AccessCodeMeeting.png" ,
    "JoinMeeting" : "https://raw.githubusercontent.com/Valenriquez/Zoom-Automation/main/ZoomAutomatizationImages/JoinMeeting.png" 
}

local_image_paths = {}
for description, url in image_info.items():
    local_path = download_image(url)
    local_image_paths[description] = local_path

join_btn = local_image_paths.get("PlusAddMeeting")
if join_btn:
    pyautogui.click(join_btn)
else:
    print("Join button not found")

meetingidbtn = pyautogui.locateCenterOnScreen(local_image_paths.get("IDMeeting"))
pyautogui.moveTo(meetingidbtn)
pyautogui.write(meeting_number_str)
time.sleep(1)

image_paths = [local_image_paths.get("Deactivate Audio"),local_image_paths.get("Deactivate Video")]
for image_path in image_paths:
    buttons = list(pyautogui.locateAllOnScreen(image_path))
    for btn in buttons:
        pyautogui.moveTo(btn)
        pyautogui.click()
        time.sleep(0.5)

joinblue_btn =  pyautogui.locateCenterOnScreen(local_image_paths.get("Join"))
pyautogui.moveTo(joinblue_btn)
pyautogui.click()
time.sleep(0.5)

passcode = pyautogui.locateCenterOnScreen(local_image_paths.get("AccessCodeMeeting"))
pyautogui.moveTo(passcode)
pyautogui.write(password_str)

joinfinal_btn =  pyautogui.locateCenterOnScreen(local_image_paths.get("JoinMeeting"))
pyautogui.moveTo(joinfinal_btn)
pyautogui.click()


