#!/usr/bin/env python
# coding: utf-8

# In[4]:


# %load opa.py
#!/usr/bin/env python

# In[1]:


"""
Created on Thu Jun 18 14:43:24 2020

@author: studentA
"""
import os
from selenium import webdriver
import tkinter as tk
from tkinter import ttk,messagebox
from datetime import date
import time
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image, ImageDraw,ImageFont,ImageTk
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from io import BytesIO
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from docx import Document
from docx.shared import Inches
from docx.shared import Cm, Pt  #加入可調整的 word 單位

class main_function():
    driver=None
    driver2=None
    name_1=None
    id_1=None
    id_2=None
    id_3=None    
    date_start=None
    date_end=None
    window=None
    code=None
    code2=None
    listbox=None
    select=None
    radioValue=None
    local_path = os.path.dirname(__file__)+'/'
    
    start_time = int(datetime.today().strftime('%Y%m%d'))-19120000
    end_time = int(datetime.today().strftime('%Y%m%d'))-19110000
    now_year = int(datetime.today().strftime('%Y'))-1911
    month_list =['','1','2','3','4','5','6','7','8','9','10','11','12']
    day_list =['','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
    year_list =['','110','109','108','107','106','105','104','103','102','101','100','99','98','97','96','95','94']
    type_list =['初發','補發','換發'] 
    locate_list =['北市','北縣','新北市','基市','桃市','桃縣','中市','中縣','南市','南縣','高市','高縣','宜縣','竹市','竹縣','苗縣','彰縣','投縣','雲縣','嘉市','嘉縣','屏縣','東縣','花縣','澎縣','連江','金門',] 
    
    def __init__(self):
        self.gui()
        
    def browser_on(self):   #開啟瀏覽器
        try:
            chrome_options = webdriver.ChromeOptions()                         #||
            chrome_options.add_argument("headless")                           #||   
            chrome_options.add_argument("--window-size=900,550")               #||
            chrome_options.add_argument("--proxy-server='direct://'");         #||
            chrome_options.add_argument("--proxy-bypass-list=*");              #||隱藏瀏覽器設定
            self.update_status('開啟瀏覽器...')
            capa = DesiredCapabilities.CHROME
            capa["pageLoadStrategy"] = "none" 
            self.driver2 = webdriver.Chrome(executable_path='chromedriver.exe',options=chrome_options,desired_capabilities=capa)
            self.driver = webdriver.Chrome(executable_path='chromedriver.exe',options=chrome_options)
            self.get_verifycode_image()
            self.update_status('開啟完成')
        except:
            messagebox.showerror(title = "未安裝驅動程式",message = '※未安裝驅動程式，請閱說明※')
            self.window.destroy()
#-------------------------------------------------------------------------------------------------------------            
    def search(self,url,site):     #查詢&截圖
        self.driver.get(url)  
        self.update_status(site+'查詢中...')
        time.sleep(0.7)
        verify_name = self.name_1.get()
        verify_id = self.id_1.get()
        verify_startt = self.date_start.get()
        verify_endt = self.date_end.get()
        
        self.driver.find_element_by_name("clnm").send_keys(verify_name)
        self.driver.find_element_by_name("idno").send_keys(verify_id)
        self.driver.find_element_by_name("sddtStart").send_keys(verify_startt)
        self.driver.find_element_by_name("sddtEnd").send_keys(verify_endt)
    
        if site == '消債事件':
            self.driver.find_elements_by_xpath('//input[@value="1"]')[0].click()
        elif site == '破產事件':
            self.driver.find_elements_by_xpath('//input[@value="2"]')[0].click()

        self.driver.find_element_by_name('Button').click()
   
        if EC.alert_is_present()(self.driver):           #錯誤處理
            self.driver.switch_to_alert().dismiss()
            self.update_status('※身份證字號錯誤，請重新輸入※') 
           
        self.update_status('查詢成功')
        _date = date.today().strftime("%Y%m%d")
        img_name = '{}{}.png'.format(site,verify_name)     #命名並儲存檔案
        
        self.driver.save_screenshot(img_name)     #截圖   
        self.update_status('已儲存{}'.format(img_name))
    
    def search_1(self):
        url='http://domestic.judicial.gov.tw/abbs/wkw/WHD9HN01.jsp'   
        self.search(url,'家事事件')
        url='http://cdcb.judicial.gov.tw/abbs/wkw/WHD9A01.jsp'   
        self.search(url,'消債事件')
        url='http://cdcb.judicial.gov.tw/abbs/wkw/WHD9A01.jsp'   
        self.search(url,'破產事件')    
    
    def search_2(self):
        verify_code2 = self.verify_code2.get()
        verify_name = self.name_1.get()
        verify_id = self.id_1.get()
#         self.driver.get('https://report.taifex.com.tw/FMS/login.html')  
        self.update_status('開戶數查詢中...')
        time.sleep(0.7)
        self.driver.find_element_by_name("username").send_keys('F008000')
        self.driver.find_element_by_name("subId").send_keys('89')
        self.driver.find_element_by_name("j_password").send_keys('Aa1234567')
        self.driver.find_element_by_name('j_captcha').send_keys(verify_code2)
        self.driver.find_elements_by_xpath('//img[@title="登入系統"]')[0].click()
        time.sleep(0.7)
        self.driver.find_element_by_id("ext-gen42").click()
        time.sleep(0.7)
        self.driver.find_element_by_id("ext-gen48").click()
        time.sleep(0.7)
        self.driver.find_element_by_id("ext-gen49").click()

        frame = self.driver.find_element_by_css_selector('iframe')
        self.driver.switch_to_frame(frame)
        time.sleep(0.7)
        self.driver.find_elements_by_xpath('/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr/td[2]/input')[0].send_keys(verify_id)
        self.driver.find_elements_by_xpath("/html/body/form/table/tbody/tr[3]/td/table/tbody/tr[3]/td/input")[0].click()
        self.update_status('查詢成功')
        
        _date = date.today().strftime("%Y%m%d")
        img_name = '{}{}.png'.format('開戶數',verify_name)     #命名並儲存檔案
        self.driver.save_screenshot(img_name)     #截圖
        self.update_status('已儲存{}'.format(img_name))
        time.sleep(0.5)
        self.update_status('違約查詢中...')
        self.driver.switch_to_default_content()
        time.sleep(0.7)
        self.driver.find_element_by_id("ext-gen45").click()
        time.sleep(0.7)
        self.driver.find_element_by_id("ext-gen46").click()
        frame = self.driver.find_element_by_css_selector('iframe')
        self.driver.switch_to_frame(frame)
        time.sleep(0.7)
        self.driver.find_elements_by_xpath("/html/body/table/tbody/tr[4]/td/input")[0].click()
        time.sleep(0.7)
        self.driver.find_elements_by_xpath('//input[@name="w_id"]')[0].send_keys(verify_id)
        self.driver.find_elements_by_xpath('//input[@value="單一身分證字號查詢"]')[0].click()
        
        _date = date.today().strftime("%Y%m%d")
        img_name = '{}{}.png'.format('違約',verify_name)     #命名並儲存檔案
        self.driver.save_screenshot(img_name)     #截圖
        self.update_status('已儲存{}'.format(img_name))

                
    def search_3(self):
        verify_name = self.name_1.get()
        self.update_status('身分證查詢中...')
        verify_id = self.id_3.get()
#         verify_date = self.issue_date.get()
        verify_day = self.combo_day.current()
        verify_month = self.combo_month.current()
        verify_year = self.combo_year.current()
        verify_site = self.combo_locate.current()
        verify_type = self.combo_type.current()
        verify_code = self.verify_code.get()

        self.driver2.find_element_by_id('idnum').send_keys(verify_id)
        selecttwy = Select(self.driver2.find_element_by_name('applyTWY'))
        selecttwy.select_by_index(verify_year)
        selectmm = Select(self.driver2.find_element_by_name('applyMM'))
        selectmm.select_by_index(verify_month)
        selectdd = Select(self.driver2.find_element_by_name('applyDD'))
        selectdd.select_by_index(verify_day)
        self.driver2.find_element_by_id('siteId').send_keys(self.locate_list[verify_site]) 
        self.driver2.find_element_by_id('applyReason').send_keys(self.type_list[verify_type]) 
        self.driver2.find_element_by_id('captchaInput_captcha-refresh').send_keys(verify_code)
        
        
        time.sleep(0.5)
        self.driver2.find_elements_by_xpath('/html/body/div[1]/div[4]/div[1]/div/form/div[4]/button[1]')[0].click()
        time.sleep(1)
        _date = date.today().strftime("%Y%m%d")
        img_name = '{}{}.png'.format('身分證',verify_name)     #命名並儲存檔案

        self.driver2.save_screenshot(img_name)     #截圖
        self.driver2.refresh()
        self.update_status('已儲存{}'.format(img_name))

        
        

#-------------------------------------------------------------------------------------------------驗證碼
    def get_verifycode_image(self):
        try:
            self.update_status('獲取驗證碼...')
            self.driver2.find_element_by_id( 'imageBlock_captcha-refresh').click()
            verify = self.driver2.find_element_by_id('captchaImage_captcha-refresh')
            self.driver2.execute_script("arguments[0].scrollIntoView();", verify)
            png = self.driver2.get_screenshot_as_png() 
            left = verify.location['x'] +250
            top = 10  #verify.location['y'] 
            elementWidth = left + verify.size['width'] -30
            elementHeight = top + verify.size['height'] -20
            picture = Image.open(BytesIO(png))
            picture = picture.crop((left, top, elementWidth, elementHeight)) 
            picture.save('勿動/verify_code.png')
            time.sleep(0.2)
            self.update_img()
            self.update_status('獲取成功')
        except:      
        
            try :
                self.driver2.get('https://www.ris.gov.tw/app/portal/3014 ')
                
                wait = WebDriverWait(self.driver2, 2)
                wait.until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[3]/div/div[1]/div/div/div/div/div/div[12]/a')))
                time.sleep(1)
                frame = self.driver2.find_element_by_css_selector('iframe')
                self.driver2.switch_to_frame(frame)
                time.sleep(1)
                self.update_status('獲取驗證碼...')
                verify = self.driver2.find_element_by_id('captchaImage_captcha-refresh')
                self.driver2.execute_script("arguments[0].scrollIntoView();", verify)
                png = self.driver2.get_screenshot_as_png() 
                left = verify.location['x'] +250
                top = 10  #verify.location['y'] 
                elementWidth = left + verify.size['width'] -30
                elementHeight = top + verify.size['height'] -20
                picture = Image.open(BytesIO(png))
                picture = picture.crop((left, top, elementWidth, elementHeight)) 
                picture.save('勿動/verify_code.png')
                time.sleep(0.2)
                self.update_img()
                self.update_status('獲取成功')
            except:
                self.get_verifycode_image()
                
    def get_verifycode2_image(self): 

        try :
            self.driver.get('https://report.taifex.com.tw/FMS/login.html')
            time.sleep(1)
            self.update_status('獲取開戶驗證碼...')
            self.driver.save_screenshot("screenshot.png")
            im = Image.open('screenshot.png') 
            im = im.crop((380, 335, 440, 370))
            im.save('勿動/verify_code2.png')
            self.update_img2()
            self.update_status('開戶驗證碼獲取成功')
        except:
            self.get_verifycode2_image()


 #-------------------------------------------------------------------------------------------------------------       
    def browser_close(self):    #關閉GUI
        try:
            self.driver2.close()
            self.driver.close()
        except:
            pass
        self.window.destroy()
        
    def gui(self):                          #介面設定
        self.window = tk.Tk()
        windowWidth = self.window.winfo_reqwidth()                                  #||
        windowHeight = self.window.winfo_reqheight()                                #||
        positionRight = int(self.window.winfo_screenwidth()/2 - windowWidth/2)+400      #||
        positionDown = int(self.window.winfo_screenheight()/2 - windowHeight/2)-300     #||
        self.window.geometry("+{}+{}".format(positionRight, positionDown))          #||設定介面出現位置
        
        self.window.title('徵信查詢系統')
        self.window.geometry('400x660')
        var = tk.StringVar()
        style = ttk.Style()
        style.configure("TButton", foreground="red", background="orange")
        
        re_img=ImageTk.PhotoImage(Image.open('勿動/refresh.png').resize((30, 30), Image.ANTIALIAS))
        img=ImageTk.PhotoImage(Image.open('勿動/白.png').resize((150, 100), Image.ANTIALIAS))
        img2=ImageTk.PhotoImage(Image.open('勿動/白.png').resize((150, 100), Image.ANTIALIAS))
#----------------------------------------------------------------------- 家事、消債查詢       
        move_x = 0
        move_y = 15
        move_ya = 90  # 身分證 button
        label1 = ttk.Label(self.window, text='姓名 :', font="微軟正黑體 15 bold").place(x =73,y=10+move_y) 
        label_id_1 = ttk.Label(self.window, text='身分證字號 :', font="微軟正黑體 15 bold").place(x =13,y=40+move_y)
        label3 = ttk.Label(self.window, text='日期範圍 :', font="微軟正黑體 15 bold").place(x =33,y=73+move_y)
        label4 = ttk.Label(self.window, text='~', font="微軟正黑體 15 bold").place(x =210,y=73+move_y)
        label10 = ttk.Label(self.window, text='※家事、消債查詢 ', font="微軟正黑體 11 bold").place(x =10,y=2)
        
        self.name_1 = ttk.Entry(self.window, font="微軟正黑體 13 ")      
        self.id_1 = ttk.Entry(self.window, font="微軟正黑體 13 ", textvariable=var)
        self.date_start = ttk.Entry(self.window, font="微軟正黑體 11 ")
        self.date_end = ttk.Entry(self.window, font="微軟正黑體 11 ")
        self.date_start.insert(0, self.start_time)
        self.date_end.insert(0, self.end_time)
        
        self.name_1.place(x =130,y=10+move_y,width=120)
        self.id_1.place(x =130,y=43+move_y,width=120)
        self.date_start.place(x =130,y=76+move_y,width=80) 
        self.date_end.place(x =230,y=76+move_y,width=80)
        
        search_btn = tk.Button(self.window, text='查詢',width=5, command= lambda: [self.search_1()], font="微軟正黑體 16 bold")
        
        separator_1 = ttk.Separator(self.window, orient="horizontal")
        separator_1.place(x=10,y=110+move_y, width=400)

#------------------------------------------------------------------ 違約、開戶數查詢       
                
        label_id_2 = ttk.Label(self.window, text='身分證字號 :', font="微軟正黑體 15 bold").place(x =13,y=146+move_y)
        label111 = ttk.Label(self.window, text='※違約、開戶數查詢 ', font="微軟正黑體 11 bold").place(x =10,y=118+move_y)
        label1_img2 = ttk.Label(self.window, text='圖形 :', font="微軟正黑體 15 bold").place(x =60,y=180+move_y)        
        label2_img2 = ttk.Label(self.window, text='驗證碼 :', font="微軟正黑體 13 bold").place(x =47,y=220+move_y)        
        self.code2 = ttk.Label(self.window,image=img2)
        self.verify_code2 = ttk.Entry(self.window, font="微軟正黑體 13 ")
            
        search_btn2 = tk.Button(self.window, text='查詢',width=5, command= lambda: [self.search_2()], font="微軟正黑體 16 bold")
        self.id_2 = ttk.Entry(self.window, font="微軟正黑體 13 ", textvariable=var).place(x =130,y=146+move_y,width=120+move_y)  
        refresh2_btn = tk.Button(self.window, text='刷新',image = re_img ,width=30,height = 30, command= lambda: self.get_verifycode2_image(), font="微軟正黑體 16 bold")

        search_btn.place(x =300,y=20+move_y)
        search_btn2.place(x =300,y=125+move_y)
        refresh2_btn.place(x =270,y=200+move_y)
        self.verify_code2.place(x =121,y=220+move_y,width=80)
        self.code2.place(x =120,y=180+move_y, width=80, height=30)        

        
        separator_2 = ttk.Separator(self.window, orient="horizontal")
        separator_2.place(x=10,y=180+move_ya, width=400)
#------------------------------------------------------------------ 身分證

        self.combo_year  = ttk.Combobox(self.window, width=3,values=self.year_list)
        self.combo_month  = ttk.Combobox(self.window, width=3,values=self.month_list)
        self.combo_day  = ttk.Combobox(self.window, width=3,values=self.day_list)
        self.combo_locate = ttk.Combobox(self.window, width=6,values=self.locate_list)
        self.combo_type = ttk.Combobox(self.window, width=5,values=self.type_list)
        self.id_3 = ttk.Entry(self.window, font="微軟正黑體 13 ", textvariable=var)
        self.verify_code = ttk.Entry(self.window, font="微軟正黑體 13 ")    
        search_btn3 = tk.Button(self.window, text='查詢',width=5, command= lambda: [self.search_3(),self.get_verifycode_image()], font="微軟正黑體 16 bold")
        search_btn4 = tk.Button(self.window, text='組合圖片',width=7, command= lambda:[self.transword()], font="微軟正黑體 11 bold")
        refresh_btn = tk.Button(self.window, text='刷新',image = re_img ,width=30,height = 30, command= lambda: self.get_verifycode_image(), font="微軟正黑體 16 bold")
        self.listbox=tk.Listbox(self.window,height=7,width=48)  


        label112 = ttk.Label(self.window, text='※身分證查詢 ', font="微軟正黑體 11 bold").place(x =10,y=188+move_ya)
        label115 = ttk.Label(self.window, text='資料起始日期： 94年12月21日 ', style="TButton", font="微軟正黑體 9 bold").place(x =110,y=188+move_ya)
        label_id_3 = ttk.Label(self.window, text='身分證字號 :', font="微軟正黑體 15 bold").place(x =13,y=215+move_ya)
        label5 = ttk.Label(self.window, text='發證日期 :', font="微軟正黑體 15 bold").place(x =33,y=245+move_ya)
        label6 = ttk.Label(self.window, text='發證地點 :', font="微軟正黑體 13 bold").place(x =29,y=275+move_ya)
        label7 = ttk.Label(self.window, text='類別 :', font="微軟正黑體 13 bold").place(x =181,y=275+move_ya)
        label8 = ttk.Label(self.window, text='圖形 :', font="微軟正黑體 15 bold").place(x =60,y=322+move_ya)        
        label113 = ttk.Label(self.window, text='驗證碼 :', font="微軟正黑體 13 bold").place(x =47,y=372+move_ya)        
        self.code = ttk.Label(self.window,image=img)
        

        
        self.combo_locate.place(x=113, y=278+move_ya)
        self.combo_year.place(x =130,y=247+move_ya)
        self.combo_month.place(x =180,y=247+move_ya)
        self.combo_day.place(x =230,y=247+move_ya)
        self.combo_type.place(x=230, y=278+move_ya)
        self.id_3.place(x =130,y=215+move_ya,width=120+move_ya)
        self.verify_code.place(x =121,y=372+move_ya,width=80)
        self.code.place(x =120,y=307+move_ya, width=140, height=60)
        search_btn3.place(x =300,y=250+move_ya)
        search_btn4.place(x =300,y=365+move_ya)
        refresh_btn.place(x =270,y=320+move_ya)
        self.listbox.place(x=30, y=410+move_ya)        
#------------------------------------------------------------------        

        self.window.resizable(0,0)
        self.window.protocol("WM_DELETE_WINDOW", self.browser_close)  #GUI按下X動作
        self.browser_on()
        self.window.mainloop()
 
    def update_img(self):
        img = ImageTk.PhotoImage(Image.open('勿動/verify_code.png').resize((125, 50), Image.ANTIALIAS))
        self.code.configure(image=img)
        self.code.image = img

    def update_img2(self):
        img2 = ImageTk.PhotoImage(Image.open('勿動/verify_code2.png').resize((80, 50), Image.ANTIALIAS))
        self.code2.configure(image=img2)
        self.code2.image = img2
        
    def update_status(self,status):                    #輸出文字到listbox
        self.listbox.insert(0, time.strftime("%H:%M:%S   ")+status) 
        self.window.update()
#---------------------------------------------------------------------------------------------轉 WORD        
    def transword(self):
        verify_name = self.name_1.get()
        img = Image.open("身分證"+verify_name+".png")
        cropped = img.crop((230,0, 850, 550))  # (left, upper, right, lower)
        cropped.save("身分證"+verify_name+".png")
        
        img = Image.open("違約"+verify_name+".png")
        cropped = img.crop((200,0, 800,170))  # (left, upper, right, lower)
        cropped.save("違約"+verify_name+".png")

        img = Image.open("開戶數"+verify_name+".png")
        cropped = img.crop((200,50, 800,300))  # (left, upper, right, lower)
        cropped.save("開戶數"+verify_name+".png")

        img = Image.open("家事事件"+verify_name+".png")
        cropped = img.crop((30,80, 900,330))  # (left, upper, right, lower)
        cropped.save("家事事件"+verify_name+".png")
        
        img = Image.open("消債事件"+verify_name+".png")
        cropped = img.crop((30,80, 900,330))  # (left, upper, right, lower)
        cropped.save("消債事件"+verify_name+".png")
        
        img = Image.open("破產事件"+verify_name+".png")
        cropped = img.crop((30,80, 900,330))  # (left, upper, right, lower)
        cropped.save("破產事件"+verify_name+".png")


#         desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
# local_path = os.path.dirname(__file__)+'/'
        file_path = os.path.dirname(os.path.abspath(__file__))
#         file_path = local_path
       # file_path = "C:\\Users\\op1\\Desktop\\Pyfile"
        pic_file_list = os.listdir(file_path)
        doc = Document()
        section = doc.sections[0]
        section.left_margin=Cm(2)
        section.right_margin=Cm(2)
        section.top_margin=Cm(2)
        section.bottom_margin=Cm(2)

        for pic_file in pic_file_list:
            if verify_name in pic_file and pic_file.endswith('.png'):
                doc.add_picture(pic_file, width=Inches(7),height=Inches(2))  # 添加圖片, 設置寬度
                doc.save(file_path + '\\' + verify_name +'.doc') # 命名存檔
                os.remove(pic_file) # 刪除原圖
            
if __name__ == '__main__':
    main_function()

