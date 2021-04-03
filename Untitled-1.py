# visual studio code를 이용하여 제작
import tkinter
import tkinter.font

from tkinter import messagebox
from tkinter import filedialog
from tkinter import Scrollbar
from tkinter import Toplevel
from tkinter import StringVar
from tkinter import ttk

import codecs
import re
import csv
import pandas as pd
import numpy as np
import os
import urllib.request
import requests
import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

from sklearn import svm, metrics
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report
from sklearn.metrics import confusion_matrix
from sklearn.metrics import precision_score
from sklearn.metrics import recall_score

window=tkinter.Tk() # tkinter 생성
window.title("Mail Analysis") # 타이틀
window.geometry("1024x720+290+30") # 너비x높이+x좌표+y좌표
window.resizable(False,False) # 상하, 좌우 크기 조절 불가
#window.configure(background='gray19')

username = StringVar()
password = StringVar()
Add_word = StringVar()

Train_Feature = []
Test_Feature = []
input_Feature = []
Input_Predict = []
data = []
lines = [] 
lastpage = 0
driver = webdriver

#모델에 필요한 데이터 불러오기
f = open('Train.csv', 'r', encoding='utf-8-sig')
Train_reader = csv.reader(f)
g = open('Test.csv', 'r', encoding='utf-8-sig')
Test_reader = csv.reader(g)
h = open('Input.csv', 'r', encoding='utf-8-sig')
Input_reader = csv.reader(h)
z = open('Filter_Word.csv', 'r', encoding='utf-8-sig')
Filter_Word = csv.reader(z)

# 로그인 창 함수
def login():
    loginView=Toplevel(window)
    loginView.title("Login")
    loginView.geometry("300x135+647+250")
    loginView.resizable(False,False)   
   
    tkinter.Label(loginView, text ="아이디").place(x=35,y=30)
    tkinter.Entry(loginView, textvariable = username).place(x=110,y=30)
    tkinter.Label(loginView, text ="비밀번호").place(x=35,y=55)
    tkinter.Entry(loginView, textvariable = password ,show='*').place(x=110,y=55)

    login_button = tkinter.Button(loginView, text="로그인", fg='white',bg='gray35',
                                activebackground='gray60', height="1", width="30", command=get) # 로그인 함수 실행
    login_button.place(x=35,y=85)

# 웹브라우저 실행해 로그인하고 메일 가져오는 함수
def get():  
    global driver
    global lastpage
    global data
    USER = username.get()
    PASS = password.get()
    page = 1

    #options = webdriver.ChromeOptions()
    #options.add_argument('headless')
    #options.add_argument("--start-maximized")
    #options.add_argument("test-type")
    #options.add_argument("--disable-gpu")
    #options.add_argument("lang=ko_KR")
    #options.add_argument('--disable-web-security')
    #options.add_argument('--allow-running-insecure-content')
    #options.add_argument('--remember-cert-error-decisions')
    #options.add_argument('--ignore-certificate-errors')
    #options.add_argument('--reduce-security-for-testing')
    #options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36')

    #driver = webdriver.Chrome('C:\chromedriver', chrome_options = options)
    driver = webdriver.Chrome('C:/chromedriver.exe')
    driver.set_window_size(1280, 900)
    driver.get('http://nid.naver.com/nidlogin.login')
    driver.execute_script(
         "document.getElementsByName('id')[0].value=\'" + USER + "\'")
    driver.execute_script(
         "document.getElementsByName('pw')[0].value=\'" + PASS + "\'")
    driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click()
    time.sleep(0.5)
    driver.get('http://mail.naver.com')
    time.sleep(0.5)

    lastpage = driver.find_element_by_css_selector("#pageSelect").get_attribute("lastpage")
    while True:
        mail_list(data) # 현재 메일 목록에서 메일 가져오는 함수

        if page == int(lastpage):
            dataframe = pd.DataFrame(data)
            dataframe.to_csv('Input.csv',header=False, index=False,encoding='utf-8-sig') 
            print('종료')
            break

        elif page % 10 == 0 and page != int(lastpage):
            page += 1
            driver.find_element_by_id('next_page').click()
            WebDriverWait(driver, 3) \
                .until(EC.presence_of_element_located((By.CSS_SELECTOR, '.mailList > li')))

        else:
            page += 1
            driver.find_element_by_xpath('//*[@id=' + str(page)+']').click()
            WebDriverWait(driver, 3) \
                .until(EC.presence_of_element_located((By.CSS_SELECTOR, '.mailList > li')))

    SVM()

# 현재 메일 목록에서 메일 가져오는 함수
def mail_list(data):
    global driver
    index = 1
    time.sleep(0.5)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
 
    WebDriverWait(driver, 10) \
        .until(EC.presence_of_element_located((By.CSS_SELECTOR, ".mailList > li")))
    mails = soup.select(".mailList > li")
 
    while (mails is None or mails == ''):
        mails = driver.find_elements_by_css_selector(".mailList > li")
    preview = driver.find_element_by_css_selector("#previewMailLayer")

    for mail in mails:
        title = mail.select_one("strong.mail_title")
        while (title is None or title == ''):
            title = mail.select_one("strong.mail_title")
        if len(title.find('span')) > 0:
            for s in title('span'):
                s.extract()
        text = re.sub('[^\u0000-\uffff]+','', str(title.text))

        name = mail.select_one(".name > a").attrs['title']
        name = re.sub('[^\u0000-\uffff]+','', str(name))
        while (name is None or name == ''):
            name = mail.select_one(".name > a").attrs['title']
            name = re.sub('[^\u0000-\uffff]+','', str(name))

        driver.find_element_by_xpath('//*[@id="list_for_view"]/ol/li['+ str(index) +']/div/div[2]/a[2]/span').click()
        index += 1

        content = preview.find_element_by_css_selector("div.pText").text
        content = re.sub('[^\u0000-\uffff]+','', str(content))
        while (content is None or content == ''):
            content = preview.find_element_by_css_selector("div.pText").text
            content = re.sub('[^\u0000-\uffff]+','', str(content))
        
        print(len(data)+1)
        row = []
        row.append(name)
        row.append(text)
        treeview.insert('', 'end',values=(len(data)+1, name, text), iid=len(data))
        row.append(content)
        data.append(row)
        preview.find_element_by_xpath('//*[@id="previewMailLayer"]/div/div[2]/button[1]').send_keys(Keys.ESCAPE)
       # print('보낸 사람 : %s \n제목 : %s \n요약 : %s\n\n\n' %(name, text, content))

# Ngram 함수
def Ngram(Reader, Feature):
 for row in Reader:
     text = re.sub('[^\uac00-\ud7a3]+','', str(row))
     a=[]
     c=[]
     for i in range(len(text) - 1):
        b=(ord(text[i])+ord(text[i+1]))%500
        a.append(b)
     for i in range(0,500):
         c.append(a.count(i))
     if (Reader == Train_reader or Reader == Test_reader):
         c.append(row[1])    
     Feature.append(c)

def SVM():
    global Input_Predict
    global data
    Ngram(Train_reader, Train_Feature)
    Ngram(Test_reader, Test_Feature)
    Ngram(Input_reader, input_Feature)

    # 데이터 프레임
    dataframe = pd.DataFrame(Train_Feature)
    Train_data = dataframe.iloc[:, 0:500]
    Train_label =  dataframe.iloc[:, 500:501]

    dataframe = pd.DataFrame(Test_Feature)
    Test_data = dataframe.iloc[:, 0:500]
    Test_label =  dataframe.iloc[:, 500:501]

    # 학습
    clf = svm.SVC()
    clf.fit(Train_data, Train_label)

    Input_Predict = clf.predict(input_Feature)
    for i, label in enumerate(Input_Predict):
        print('보낸 사람 : %s \n제목 : %s \n요약 : %s\n결과 : %s\n\n' %(data[i][0], data[i][1], data[i][2], label))

    # 학습된 스크레이핑 메일에서 필터링 단어 있으면 라벨 Non-spam 처리
    h.seek(0)
    for index, y in enumerate(Input_reader):
         z.seek(0)
         for x in Filter_Word:
              word = x[0]
              text = ",".join(y)
              if(text.find(word) != -1):
                 Input_Predict[index] = 'Non-spam'

    for i, label in enumerate(Input_Predict):
        print('보낸 사람 : %s \n제목 : %s \n요약 : %s\n결과 : %s\n\n' %(data[i][0], data[i][1], data[i][2], label))         

    Test_Predict = clf.predict(Test_data)
    confusion = confusion_matrix(Test_Predict,Test_label, labels=("spam", "Non-spam"))
    
    # 결과
    print("오차 행렬:\n{}".format(confusion))
    print("정밀도(precision) = ", precision_score(Test_label,Test_Predict, pos_label="spam"))
    print("정답률(accuracy) = ", metrics.accuracy_score(Test_label,Test_Predict))
    print("재현율(Recall) = ", recall_score(Test_label,Test_Predict, average="binary", pos_label="spam"))
    report = classification_report(Test_label,Test_Predict, digits=4)
    print(report)
    

def retrain():
    # 일단 아무거나 써논거
    loginView=Toplevel(window)
     
def filtering():
    global lines
    lines = []
    
    treeview.destroy()
    get_button.destroy()
    filter_button.destroy()
    block_button.destroy()
    analysis_button.destroy()

    entry.place(x=365,y=500)
    listbox.place(y=150, x=365)
    if(listbox.size() > 0): # 리스트박스가 이미 존재하고 있으면
        listbox.delete(0, listbox.size())
    z = open('Filter_Word.csv', 'r', encoding='utf-8-sig',  newline='')
    Filter_Word = csv.reader(z) 
    for i, row in enumerate(Filter_Word):
        lines.append(row[0])
        listbox.insert(i, row[0])

    add_button.place(x=500,y=550)
    delete_button.place(x=580,y=550)

def add():
    text = Add_word.get()
    entry.delete(0, 'end')
    if (text not in lines and text != ''):
     lines.append(text)
     listbox.insert(tkinter.END, text)

     z = open('Filter_Word.csv', 'a', encoding='utf-8-sig',  newline='')
     Filter_Word = csv.writer(z) 
     Filter_Word.writerow([text])
     z.close()

def delete():
    sel = listbox.curselection()

    z = open('Filter_Word.csv', 'wr', encoding='utf-8-sig', newline='')
    Filter_Word = csv.reader(z) 
    for i,line in enumerate(Filter_Word):
        if i not in sel:
            Filter_Word.writerow([line])

    z.close()            
    for index in reversed(sel):
         listbox.delete(index)

def back():
    listbox.destroy()
    add_button.destroy()
    delete_button.destroy()
    back_button.destroy()
    
    back_button.place(x=500, y=600)

def analysis():
    global Input_Predict
    # 일단 아무거나 써논거
    for i, label in enumerate(Input_Predict):
        treeview.set(i, column="result", value=label)

#스팸 처리
def block():
 global driver
 global lastpage
 global Input_Predict
 Lastpage_spam_check = "" # 마지막 페이지 메일이 모두 스팸인지 검사하는 변수

 if(int(len(Input_Predict) % 10) == 0):
     pages = int(len(Input_Predict) / 10) - 1

 else:  # len(Input_Predict) % 10) != 0):
     pages = int(len(Input_Predict) / 10)

 for page in range(pages, -1, -1):
     indexs = 10

     if(page == int(lastpage) - 1): # 마지막 페이지 메일 개수 검사
         indexs = len(driver.find_elements_by_css_selector(".mailList > li"))
         for index in range(indexs, 0, -1):
             if(Input_Predict[((page * 10) + index) - 1] == 'Non-spam'):  # 마지막 페이지에서 스팸 아닌 메일 찾음
                 Lastpage_spam_check = "No" # 있으면 직접 페이지 넘겨야함 
         #len(Input_Predict) - page * indexs

     for index in range(indexs, 0, -1):
         if(Input_Predict[((page * 10) + index) - 1] == 'spam'):  # 스팸이면
             WebDriverWait(driver, 10) \
             .until(EC.presence_of_element_located((By.CSS_SELECTOR, 'ol > li:nth-of-type(' + str(index) + ') input')))
             driver.find_element_by_css_selector('ol > li:nth-of-type(' + str(index) + ') input').click()

     if(driver.find_element_by_xpath('//*[@id="listBtnMenu"]/div[1]/span[2]/button[4]').is_enabled()): # 스팸이 있으면
         driver.find_element_by_xpath('//*[@id="listBtnMenu"]/div[1]/span[2]/button[4]').click()
         WebDriverWait(driver, 10) \
         .until(EC.presence_of_element_located((By.ID, 'submitDeleteSpamBtn')))
         driver.find_element_by_id('submitDeleteSpamBtn').click()

     if(page > 0):
         if(page == int(lastpage) - 1): # 마지막 페이지일때
             if(Lastpage_spam_check == "No"): # 정상이 하나라도 있으면
                 driver.find_element_by_css_selector('.paging_numbers > a:nth-of-type(' + str(page) + ')').click() # 페이지 직접 넘김
         else: # 마지막 페이지 아니면
             driver.find_element_by_css_selector('.paging_numbers > a:nth-of-type(' + str(page) + ')').click()                
         time.sleep(0.5)

get_button = tkinter.Button(window, text="메일 불러오기", fg='white',bg='gray35', font = 'TkDefaultFont 15',
                                activebackground='gray60', height="1", width="15",command=login)
get_button.place(x=40,y=670)

filter_button = tkinter.Button(window, text="단어 필터링", fg='white',bg='gray35', font = 'TkDefaultFont 15',
                             activebackground='gray60', height="1", width="15",command=filtering)
filter_button.place(x=250,y=670)

analysis_button = tkinter.Button(window, text="분석", fg='white',bg='gray35', font = 'TkDefaultFont 15',
                             activebackground='gray60', height="1", width="5",command=analysis)
analysis_button.place(x=830,y=670)

block_button = tkinter.Button(window, text="차단", fg='white',bg='gray35', font = 'TkDefaultFont 15',
                             activebackground='gray60', height="1", width="5",command=block)
block_button.place(x=915,y=670)

listbox = tkinter.Listbox(window, selectmode='extended', height=20, width=43)

entry=tkinter.Entry(window, textvariable = Add_word, width=41)

add_button = tkinter.Button(window, text="추가", fg='white',bg='gray35', font = 'TkDefaultFont 12',
                             activebackground='gray60', height="1", width="5",command=add)

delete_button = tkinter.Button(window, text="삭제", fg='white',bg='gray35', font = 'TkDefaultFont 12',
                             activebackground='gray60', height="1", width="5",command=delete)

back_button = tkinter.Button(window, text="RE", fg='white',bg='gray35', font = 'TkDefaultFont 12',
                             activebackground='gray60', height="1", width="1",command=back)

treeview=ttk.Treeview(window, columns=["num","sender","title", "result"], height="30")
treeview.place(x=40, y=30)


treeview.column("num", width=37, anchor="center")
treeview.heading("num", anchor="center")

treeview.column("sender", width=235, anchor="w")
treeview.heading("sender", text="보낸사람",anchor="center")

treeview.column("title", width=535, anchor="w")
treeview.heading("title", text="제목", anchor="center")

treeview.column("result", width=130, anchor="w")
treeview.heading("result", text="결과", anchor="center")

treeview['show']='headings'
window.mainloop()