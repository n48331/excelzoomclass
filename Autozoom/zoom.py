import time
import pyautogui as pi
import openpyxl
import schedule 
import subprocess

book= openpyxl.load_workbook('data.xlsx')
sheet=book.active

l='C:/Users/name/AppData/Roaming/Zoom/bin/Zoom.exe'#your zoom.exe file location.change it completely


def zoom1():
    z=subprocess.call([l])#your zoom.exe file location
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet['A1'].value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet['B1'].value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

def zoom2():
    z=subprocess.call([l])#your zoom.exe file location
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet['A2'].value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet['B2'].value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

def zoom3():
    z=subprocess.call([l])#your zoom.exe file location
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet['A3'].value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet['B3'].value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)

def zoom4():
    z=subprocess.call([l])
    time.sleep(1)
    r=pi.locateCenterOnScreen('2.png')
    print(r)
    pi.click(r)
    time.sleep(1)
    pi.write(str(sheet['A3'].value),interval=0.05)
    time.sleep(1)
    j=pi.locateCenterOnScreen('join.png')
    pi.click(j)
    time.sleep(2)
    pi.write(str(sheet['B3'].value),interval=0.05)
    time.sleep(1)
    jm=pi.locateCenterOnScreen('join2.png')
    pi.click(jm)
    time.sleep(3)
    
print("Started Schedules")
schedule.every().day.at(str(sheet['C1'].value)).do(zoom1)
schedule.every().day.at(str(sheet['C2'].value)).do(zoom2)
schedule.every().day.at(str(sheet['C3'].value)).do(zoom3)
schedule.every().day.at(str(sheet['C4'].value)).do(zoom4)
while True: 
	schedule.run_pending()
	time.sleep(1) 
schedule.CancelJob()