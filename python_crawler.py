import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
import time
import os


class Gnr_GUI:
    url = 'https://judgment.judicial.gov.tw/FJUD/Default_AD.aspx'
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'plugins.always_open_pdf_externally': True,
        })
    
    global driver
    driver = webdriver.Chrome(options=chrome_options)   
    driver.get(url)    

    global gui
    gui = tk.Tk()
    gui.title('批量下載判決書工具')
    window_width = gui.winfo_screenwidth()    # 取得螢幕寬度
    window_height = gui.winfo_screenheight()  # 取得螢幕高度
    gui.configure(bg='#323232')
    width = 800
    height = 450
    left = int((window_width - width)/2)       # 計算左上 x 座標
    top = int((window_height - height)/2)      # 計算左上 y 座標
    gui.geometry(f'{width}x{height}+{left}+{top}')
    gui.resizable(False, False)

    
    def __init__(self):
        self.G_IB()
        self.G_CB()
        self.G_EB()
        self.G_TB()
        self.G_SB()
        self.gui_Btn()
    
    def G_IB(self):      #法院選擇 generate_itembox
        frame = tk.Frame(gui, width=24)        
        frame.place(x=30,y=30)

        scrollbar = tk.Scrollbar(frame)        
        scrollbar.pack(side='right', fill='y')  

        global menu
        menu = tk.StringVar()
        menu.set(('所有法院', '憲法法庭', '司法院刑事補償法庭', '司法院－訴願決定', '最高法院', '最高行政法院(含改制前行政法院)', 
          '懲戒法院－懲戒法庭', '懲戒法院－職務法庭', '臺灣高等法院', '臺灣高等法院－訴願決定', 
          '臺北高等行政法院 高等庭(含改制前臺北高等行政法院)', '臺北高等行政法院 地方庭', 
          '臺中高等行政法院 高等庭(含改制前臺中高等行政法院)', '臺中高等行政法院 地方庭', 
          '高雄高等行政法院 高等庭(含改制前高雄高等行政法院)', '高雄高等行政法院 地方庭', 
          '智慧財產及商業法院', '臺灣高等法院 臺中分院', '臺灣高等法院 臺南分院', '臺灣高等法院 高雄分院', 
          '臺灣高等法院 花蓮分院', '臺灣臺北地方法院', '臺灣士林地方法院', '臺灣新北地方法院', '臺灣宜蘭地方法院', 
          '臺灣基隆地方法院', '臺灣桃園地方法院', '臺灣新竹地方法院', '臺灣苗栗地方法院', '臺灣臺中地方法院', 
          '臺灣彰化地方法院', '臺灣南投地方法院', '臺灣雲林地方法院', '臺灣嘉義地方法院', '臺灣臺南地方法院', 
          '臺灣高雄地方法院', '臺灣橋頭地方法院', '臺灣花蓮地方法院', '臺灣臺東地方法院', '臺灣澎湖地方法院', 
          '福建高等法院金門分院', '福建金門地方法院', '福建連江地方法院', '臺灣高雄少年及家事法院'))        #：多嗎？ 我覺得蠻多的

        global nick
        nick = {'所有法院':'ALL', '憲法法庭':'JCC', '司法院刑事補償法庭':'TPC', '司法院－訴願決定':'TPU', '最高法院':'TPS',
         '最高行政法院(含改制前行政法院)':'TPA', '懲戒法院－懲戒法庭':'TPP', '懲戒法院－職務法庭':'TPJ', '臺灣高等法院':'TPH', 
         '臺灣高等法院－訴願決定':'001', '臺北高等行政法院 高等庭(含改制前臺北高等行政法院)':'TPB', '臺北高等行政法院 地方庭':'TPT', 
         '臺中高等行政法院 高等庭(含改制前臺中高等行政法院)':'TCB', '臺中高等行政法院 地方庭':'TCT', 
         '高雄高等行政法院 高等庭(含改制前高雄高等行政法院)':'KSB', '高雄高等行政法院 地方庭':'KST', 
         '智慧財產及商業法院':'IPC', '臺灣高等法院 臺中分院':'TCH', '臺灣高等法院 臺南分院':'TNH', '臺灣高等法院 高雄分院':'KSH', 
         '臺灣高等法院 花蓮分院':'HLH', '臺灣臺北地方法院':'TPD', '臺灣士林地方法院':'SLD', '臺灣新北地方法院':'PCD', 
         '臺灣宜蘭地方法院':'ILD', '臺灣基隆地方法院':'KLD', '臺灣桃園地方法院':'TYD', '臺灣新竹地方法院':'SCD', 
         '臺灣苗栗地方法院':'MLD', '臺灣臺中地方法院':'TCD', '臺灣彰化地方法院':'CHD', '臺灣南投地方法院':'NTD', 
         '臺灣雲林地方法院':'ULD', '臺灣嘉義地方法院':'CYD', '臺灣臺南地方法院':'TND', '臺灣高雄地方法院':'KSD', 
         '臺灣橋頭地方法院':'CTD', '臺灣花蓮地方法院':'HLD', '臺灣臺東地方法院':'TTD', '臺灣澎湖地方法院':'PHD', 
         '福建高等法院金門分院':'KMH', '福建金門地方法院':'KMD', '福建連江地方法院':'LCD', '臺灣高雄少年及家事法院':'KSY'}

        global listbox
        listbox = tk.Listbox(frame, listvariable=menu, selectmode='extended', height=18, width=24, yscrollcommand = scrollbar.set)
        listbox.pack(side='left', fill='y')    
        scrollbar.config(command = listbox.yview)

    def G_CB(self):        #案件類別多選 checkbox
        cx, cy = 280, 30
        global varC, varV, varM, varA, varP
        varC = tk.IntVar()
        varV = tk.IntVar()
        varM = tk.IntVar()
        varA = tk.IntVar()
        varP = tk.IntVar()
        c_box_C = tk.Checkbutton(gui, text='憲法', variable=varC).place(x=cx+80,y=cy+10)
        c_box_V = tk.Checkbutton(gui, text='民事', variable=varV).place(x=cx+140,y=cy+10)
        c_box_M = tk.Checkbutton(gui, text='刑事', variable=varM).place(x=cx+200,y=cy+10)
        c_box_A = tk.Checkbutton(gui, text='行政', variable=varA).place(x=cx+260,y=cy+10)
        c_box_P = tk.Checkbutton(gui, text='懲戒', variable=varP).place(x=cx+320,y=cy+10)

        tk.Label(gui, text='案件類別',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=cx,y=cy) 

    def G_EB(self):       #裁判字號輸入 entrybox
        wx, wy = 280, 70
        global year, case, no, no_end
        year = tk.StringVar()   
        case = tk.StringVar()
        no = tk.StringVar()
        no_end = tk.StringVar()
        jud_year = tk.Entry(gui, width=3, textvariable=year).place(x=wx+80,y=wy+10)
        jud_case = tk.Entry(gui, width=10, textvariable=case).place(x=wx+160,y=wy+10)
        jud_no = tk.Entry(gui, width=5, textvariable=no).place(x=wx+310,y=wy+10)
        jud_no_end = tk.Entry(gui, width=5, textvariable=no_end).place(x=wx+395,y=wy+10)

        tk.Label(gui, text='年度', font=('Arial',16), fg='#fff').place(x=wx+120,y=wy+10)
        tk.Label(gui, text='字 第', font=('Arial',16), fg='#fff').place(x=wx+265,y=wy+10)
        tk.Label(gui, text='—', font=('Arial',16), fg='#fff').place(x=wx+370,y=wy+10)
        tk.Label(gui, text='號', font=('Arial',16), fg='#fff').place(x=wx+455,y=wy+10)

        tk.Label(gui, text='裁判字號',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=wx,y=wy)

    def G_TB(self):        #裁判期間輸入 timebox
        tx, ty = 280, 110
        global dy1, dm1, dd1, dy2, dm2, dd2
        dy1 = tk.StringVar()
        dm1 = tk.StringVar()
        dd1 = tk.StringVar()
        dy2 = tk.StringVar()
        dm2 = tk.StringVar()
        dd2 = tk.StringVar()
        jud_dy1 = tk.Entry(gui, width=3, textvariable=dy1).place(x=tx+120,y=ty+10)
        jud_dm1 = tk.Entry(gui, width=2, textvariable=dm1).place(x=tx+180,y=ty+10)
        jud_dd1 = tk.Entry(gui, width=2, textvariable=dd1).place(x=tx+230,y=ty+10)
        jud_dy2 = tk.Entry(gui, width=3, textvariable=dy2).place(x=tx+310,y=ty+10)
        jud_dm2 = tk.Entry(gui, width=2, textvariable=dm2).place(x=tx+370,y=ty+10)
        jud_dd2 = tk.Entry(gui, width=2, textvariable=dd2).place(x=tx+420,y=ty+10)

        tk.Label(gui, text='民國', font=('Arial',16), fg='#fff').place(x=tx+80,y=ty+10)
        tk.Label(gui, text='年', font=('Arial',16), fg='#fff').place(x=tx+160,y=ty+10)
        tk.Label(gui, text='月', font=('Arial',16), fg='#fff').place(x=tx+210,y=ty+10)
        tk.Label(gui, text='日  至', font=('Arial',16), fg='#fff').place(x=tx+260,y=ty+10)
        tk.Label(gui, text='年', font=('Arial',16), fg='#fff').place(x=tx+350,y=ty+10)
        tk.Label(gui, text='月', font=('Arial',16), fg='#fff').place(x=tx+400,y=ty+10)
        tk.Label(gui, text='日', font=('Arial',16), fg='#fff').place(x=tx+450,y=ty+10)

        tk.Label(gui, text='裁判期間',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=tx,y=ty)

    def G_SB(self):      #案由主文內容輸入 searchbox
        ex, ey = 280, 150
        global title, jmain, kw
        title = tk.StringVar()
        jmain = tk.StringVar()
        kw = tk.StringVar()
        jud_title = tk.Entry(gui, width=24, textvariable=title).place(x=ex+80,y=ey+10)
        jud_jmain = tk.Entry(gui, width=40, textvariable=jmain).place(x=ex+80,y=ey+50)
        jud_kw = tk.Entry(gui, width=40, textvariable=kw).place(x=ex+80,y=ey+90)

        tk.Label(gui, text='裁判案由',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=ex,y=ey)

        tk.Label(gui, text='裁判主文',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=ex,y=ey+40)

        tk.Label(gui, text='全文內容',
                 font=('Arial',16,'bold'),
                 bg='#4d4d4d', fg='#e6e6e6',
                 padx='5', pady='10',
                 relief='raised').place(x=ex,y=ey+80)

    def gui_Btn(self):
        global Btn
        Btn1 = tk.Button(gui, text='送出', command=self.confirm)
        Btn1.place(x=500,y=300)

    def confirm(self):
        global new
        new = tk.Toplevel(master=gui)
        new.title('下載頁面')
        new.geometry(f'{self.width}x{self.height}+{self.left}+{self.top}')
        new.resizable(False, False)
        new.configure(bg='#323232')

        global wd1
        global wd2
        wd1 = tk.StringVar()
        wd2 = tk.StringVar()
        wd1.set('等待確認中...')
        wd2.set('')
        dl1 = tk.Label(new, textvariable=wd1, font=('Arial',24), fg='#fff')
        dl1.place(x=140,y=150)
        dl2 = tk.Label(new, textvariable=wd2, font=('Arial',24), fg='#fff')
        dl2.place(x=500,y=150)

        global bar
        bar = ttk.Progressbar(new, length=560)
        bar.place(x=120,y=200)
        bar['value'] = 0

        global Btn2
        global Btn3
        Btn2 = tk.Button(new, text='按我下載', command=Next)
        Btn2.place(x=180,y=300)
        Btn3 = tk.Button(new, text='結束程式', command=self.quit, state='disabled')
        Btn3.place(x=540,y=300)

    def mainloop(self):
        gui.mainloop()

    def quit(self):
        gui.destroy()



class Next:
    def __init__(self):
        Btn2['state'] = 'disabled'

        self.itembox()
        self.checkbox()
        self.entrybox()
        self.timebox()
        self.searchbox()
        self.submit()

        crawler()

    def itembox(self):
        box = []
        a = listbox.curselection()
        for i in a:
            box.append(nick.get(listbox.get(i)))

        for j in box:
            if j =='ALL':
                SELECT_BOX = driver.find_element(By.XPATH, '//option[@selected="selected"]')
                SELECT_BOX.click()
                break
            else:
                SELECT_BOX = driver.find_element(By.XPATH, f'//option[@value="{j}"]')
                SELECT_BOX.click()

    def checkbox(self):
        box = [varC.get(), varV.get(), varM.get(), varA.get(), varP.get()]
        CHECK_BOX = driver.find_elements(By.XPATH, '//input[@name="jud_sys"]')
        for i in range(5):
            a = box[i]
            if a == 1:
                CHECK_BOX[i].click()
            else:
                continue
    
    def entrybox(self):
        box = [year.get(), case.get(), no.get(), no_end.get()]
        boxname = ['jud_year', 'jud_case', 'jud_no', 'jud_no_end']
        for i in range(4):
            ENTRY_BOX = driver.find_element(By.XPATH, f'//input[@name="{boxname[i]}"]')
            a = box[i]
            ENTRY_BOX.send_keys(a)

    def timebox(self):
        box = [dy1.get(), dm1.get(), dd1.get(), dy2.get(), dm2.get(), dd2.get()]
        boxname = ['dy1', 'dm1', 'dd1', 'dy2', 'dm2', 'dd2']
        for i in range(6):
            TIME_BOX = driver.find_element(By.XPATH, f'//input[@name="{boxname[i]}"]')
            a = box[i]
            TIME_BOX.send_keys(a)
    
    def searchbox(self):
        box = [title.get(), jmain.get(), kw.get()]
        boxname = ['jud_title', 'jud_jmain', 'jud_kw']
        for i in range(3):
            SEARCH_BOX = driver.find_element(By.XPATH, f'//input[@name="{boxname[i]}"]')
            a = box[i]
            SEARCH_BOX.send_keys(a)

    def submit(self):
        submit_Btn = driver.find_element(By.XPATH, '//input[@type="submit"]')
        submit_Btn.click()
 
def jud_error():
    jud_error = 'https://judgment.judicial.gov.tw/Errorpage.aspx?aspxerrorpath=/FJUD/Default_AD.aspx'
    if driver.current_url == jud_error:
        wd1.set('司法院系統忙碌中，請稍後再試')
        Btn3['state'] = 'normal'
        new.update()

def crawler():
    try:
        number = int(driver.find_element(By.CLASS_NAME, 'badge').text)
        if number >= 500:
            wd1.set(f'搜索到{number}個結果，僅下載前500份資料')
            new.update()
            number = 500
        else:
            wd1.set(f'搜索到{number}個結果')
            new.update()

        
        aim_web = driver.find_element(By.XPATH, '//iframe').get_attribute('src')
        driver.get(aim_web)

        page = number//20 + 1
        download_times = 0
    
        for i in range(1, page+1):
            original_window = driver.current_window_handle
            try:
                judgment = driver.find_elements(By.XPATH, '//a[@id="hlTitle"]')     # 判決資料
            except:
                jud_error()

            wd1.set(f'正在下載第{i}/{page}頁')
            bar['value'] = 0
            bar['maximum'] = len(judgment)
            new.update()

            for j in judgment :
                link = j.get_attribute('href')
                name = j.text   #判決字號
                driver.switch_to.new_window('tab')
                try :
                    driver.get(link)
                except:
                    jud_error()

                pdf_Btn = driver.find_element(By.ID, 'hlExportPDF').get_attribute('href')
                pdf = requests.get(pdf_Btn)
                with open(f'judgment/{name}.pdf', 'wb') as f:
                    f.write(pdf.content)

                download_times += 1
                bar['value'] += 1
                wd2.set(f'成功下載{download_times}/{number}份')
                new.update()
                driver.close()
                driver.switch_to.window(original_window)
            try:
                nextpage = driver.find_element(By.ID, "hlNext")
                nextpage.click()
            except:
                pass

        wd1.set(f'已完成下載')
        Btn3['state'] = 'normal'
        new.update()

    except:
        jud_error()


if __name__ == '__main__':
    try:
        os.mkdir('judgment') # 下載的資料會在這
    except:
        pass

    Gnr_GUI()
    Gnr_GUI().mainloop()  
  