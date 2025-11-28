#import java.util.concurrent.TimeUnit
#from xxlimited import Str
import six
import packaging
import packaging.version
import packaging.specifiers
import packaging.requirements
from numpy import var
import openpyxl
from asyncio import sleep, wait, wait_for
import time
import timeit
#import cv2
#from tqdm import tqdm
#import matplotlib.pyplot as plt


from selenium.common.exceptions import TimeoutException
from PIL import Image
#import pytesseract
import os
import numpy as np
#from websb_manager.chrome import ChromesbManager
from openpyxl import Workbook
import datetime
from selenium.common.exceptions import ElementNotInteractableException
#import execjs
#from selenium.websb.chrome.options import Options
#pytesseract.pytesseract.tesseract_cmd = r'C:\Users\duong.ns\AppData\Local\Tesseract-OCR\tesseract.exe'
from seleniumbase import SB
OP_DIR = r'./web'
Cap_DIR = r'./Captcha'
#THAY DOI THONG TIN ÐANG NH?P
username = "2500267686"
password = "2500267686"
soluong = input("Nhập so luong ma dinh danh: ")
S = lambda X: sb.execute_script('return document.body.parentNode.scroll'+X)
#Sumdonghang = int(input()) - 1

#40options.add_argument('--headless')
#options.add_argument('--disable-gpu')  # Tắt GPU acceleration
#options.add_argument('--no-sandbox')  # Tắt chế độ sandbox

#options.add_argument("--start-maximized")
#options.add_argument('--blink-settings=imagesEnabled=false')  # Tắt hiển thị hình ảnh
#sb = websb.Chrome(options=options)
with SB() as sb:
    sb.get("https://pus.customs.gov.vn/faces/Login")
    # find username/email field and send the username itself to the input field
    #time.sleep(5)
    tendn = sb.find_element("#pt1\:it1\:\:content")
    tendn.send_keys(username)
    passw = sb.find_element("#pt1\:it2\:\:content")
    passw.send_keys(password)
    #nut dang nhap: /html/body/div[2]/form/div/div[4]/div[1]/span/div[1]/table/tbody/tr[6]/td[2]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/div
    sb.find_element("#pt1\:rsoLoginType\:_1").click()
    ###
    #sb.set_window_size(S('Width'),S('Height'))
    #sb.find_element_by_tag_name('body').screenshot(f'{OP_DIR}\{1}.png')
    #iname = f'{OP_DIR}\{1}.png'
    #img = cv2.imread(iname)
    #cropped_image = img[455:497,470:615]
    #cropped_image = img[342:399,510:590]
    #cv2.imshow("Cropped Image", img)
    #cropped_image.save('./test.png')
    #cv2.imshow("cropped", cropped_image)

    ################################ Doc Capcha

    #cv2.imwrite("./captcha/new_image.jpg", cropped_image)
    #image = cv2.imread("./captcha/new_image.jpg")
    #gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    #gray = cv2.threshold(gray, 0, 255,
            #cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1] 

    #filename = "{}.png".format(os.getpid())

    #cv2.imwrite(filename, gray)
    #cv2.imwrite(filename)
    # Load ảnh và apply nhận dạng bằng Tesseract OCR
    #custom_config = r'--oem 3 --psm 6'
    maxacthuc = input("Nhap capcha: ")
    #maxacthuc = pytesseract.image_to_string(Image.open("./captcha/new_image.jpg"),lang='eng')
    #maxacthuc = pytesseract.image_to_string(image, lang='eng')
    #maxacthuc = maxacthuc.split(" ")
    #cv2.imshow("Image", image)
    #cv2.imshow("Output", gray)
    #cv2.waitKey(0)
    # Xóa ảnh tạm sau khi nhận dạng
    #os.remove(filename)
    #sb.find_element(By.ID,"pt1:it42::content").send_keys(maxacthuc)
    #print (maxacthuc)
    sb.find_element("#pt1\:it42\:\:content").send_keys(maxacthuc)
    # click login button
    #sb.find_element(By.XPATH,"/html/body/div[1]/form/div/div[4]/div[1]/span/div[1]/table/tbody/tr[6]/td[2]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/div").click()
    sb.find_element("#pt1\:cbLogin").click()
    element1 = sb.find_element("#pt1\:dc7\:dinhDanhhangHoa > div > table > tbody > tr > td.x18i > a")

    # Click vào liên kết
    element1.click()
    nutcapmoi = sb.find_element("#pt1\:b1")
    sb.execute_script("arguments[0].click();", nutcapmoi)
    #nutcapmoi.click()

    # Chon DN XNK
    # Wait for the select element to be visible and select "DN XNK" option
    dropdown_selector = "#pt1\\:soc2\\:\\:content"
    sb.select_option_by_text(dropdown_selector, "Xuất khẩu")
    #TK XUAT
    dnxk = "#pt1\:soc1\:\:content"
    sb.select_option_by_text(dnxk, "DN XNK")
    #TK NHAP
    #select.select_by_value("0")
    sb.find_element("#pt1\:it5\:\:content").send_keys(username)
    time.sleep(1)
    select_element2 = sb.find_element("#pt1\:cb3")
    #select_element2.click()
    sb.execute_script("arguments[0].click();", select_element2)
    #sb.find_element(By.ID,"pt1:cb3").click()
    #COPY so dinh danh vao excel
    span_element = sb.find_element("#pt1\:it11\:\:content")
    text = span_element.text
    #Dong
    dong = sb.find_element("#pt1\:b4")
    sb.execute_script("arguments[0].click();", dong)

    #/html/body/div[1]/form/div[2]/div[2]/div[1]/div[1]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[2]/table[2]/tbody/tr/td[1]
    # Create a new Excel workbook and write the text to a cell
    #workbook = Workbook()
    workbook = openpyxl.load_workbook(r"F:\desktop\So dinh danh\output.xlsx")
    sheet = workbook.active
    now = datetime.date.today()
    last_row = 2

    #last_row = sheet["A1"].value
    sheet["C" + str(last_row + 1)] = text
    sheet["D" + str(last_row + 1)] = "EXPORT"
    sheet["B" + str(last_row + 1)] = "2500267686"
    sheet["E" + str(last_row + 1)] = now
    #sheet["D" + str(last_row + 1)] = "Ngày update"
    print (text)
    workbook.save
    workbook.save(r"F:\desktop\So dinh danh\output.xlsx")
    #p = input()
    #sheet["F" + str(last_row + 1)] = now
    for i in range(2, int(soluong)+1):
        nutcapmoi = sb.find_element("#pt1\:b1")
        time.sleep(1)
        #nutcapmoi.click()
        sb.execute_script("arguments[0].click();", nutcapmoi)
        
        nutcapmoi1 = select_element2 = sb.find_element("#pt1\:cb3",timeout = 50)
        
        #time.sleep(0.5)
        #select_element2.click()d((By.LINK_TEXT, "Cấp mới")))
        #wait = WebsbWait(sb, 10)
        #nutcapmoi1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Cấp mới']")))
        sb.execute_script("arguments[0].click();", nutcapmoi1)
        span_element = sb.find_element("#pt1\:it11\:\:content",timeout = 50)
        time.sleep(0.5)
        text = span_element.text
        sheet["C" + str(last_row + i)] = text
        sheet["D" + str(last_row + i)] = "EXPORT"
        sheet["B" + str(last_row + i)] = "2500267686"
        sheet["E" + str(last_row + i)] = now
        print (text)
        dong = sb.find_element("#pt1\:b4")
        sb.execute_script("arguments[0].click();", dong)
        workbook.save
    # Lưu workbook vào một tệp
    workbook.save(r"F:\desktop\So dinh danh\output.xlsx")

# Đóng trình duyệt
#sb.quit()
#sb.find_element(By.ID,"ctl00_Header1_lbtnLogin").click()

