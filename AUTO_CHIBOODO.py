import time
import openpyxl
import pyautogui
#
#
#xls 파일은 인식불가, xlsx 파일로 변환(혹은 새로 만들어서) 사용
#실행 후 취부도 레이아웃으로 화면 전환 필수
#실행 전 DOT을 클립보드에 복사해놓고 시작(바로바로 Ctrl+V 할 수 있도록)
#
#
#################################부품번호 추출#################################

wb = openpyxl.load_workbook(r"C:\Users\user\Documents\AutoPython\4M_CHIBOO.xlsx") #PCB_BOM 엑셀파일 경로
ws = wb["Sheet1"]

#Value, Name열번호 추출
flag = 0 #이중루프를 빠져나오기 위한 변수
for i in range(1,10):
    for row in ws[i]:
        index = row.value
        if index == "Value":
            Val = row.column - 1
        if index == "Name":
            Name = row.column - 1
        if 'Val' in locals() and 'Name' in locals():
            flag = 1
            break
    if flag:
        break
#
# print(Val, Name)
#
# DOT과 DNI에 해당하는 부품번호 추출
DOT = [] # 점 찍어야 하는 부품 리스트
DNI = [] # DNI 부품 리스트
for row in ws.rows:
    DOT_Name = row[Name].value

    if DOT_Name is not None:
        if row[Val].value == "DNI":
            DNI.append(row[Name].value)
            print(row[Name].value)

        elif DOT_Name[0:2] != "SM" and DOT_Name[0:5] != "U4900":
            if DOT_Name[0] in ("U","D", "S", "Y") or DOT_Name[0:2] in ("FL", "SW"):
                DOT_Name_1 = DOT_Name + ".1" #첫 번째 포트번호 추가
                DOT.append(DOT_Name_1)
                print(DOT_Name_1)


#
# print(DOT)
# print(DNI)
#################################매크로 시작#################################

time.sleep(3) #실행 후 3초뒤 매크로 시작

# DOT 매크로
for i in DOT:
    pyautogui.typewrite("s")
    pyautogui.hotkey('space')
    pyautogui.typewrite(i)
    pyautogui.hotkey('enter')
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.click()
    time.sleep(0.4)

time.sleep(3)

# DNI 매크로
for j in DNI:
    pyautogui.typewrite("s")
    pyautogui.hotkey('space')
    pyautogui.typewrite(j)
    pyautogui.hotkey('enter')
    time.sleep(0.2)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.hotkey('delete')
    time.sleep(0.7)


