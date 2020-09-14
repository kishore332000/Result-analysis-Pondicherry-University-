# -*- coding: utf-8 -*-
"""
Created on Fri Feb  7 10:19:11 2020

@author: kishore kannan S,Poovanan G
"""
import xlsxwriter
from selenium import webdriver
import pandas as pd
from tkinter import Tk,Entry,Label,Button
sem=['A','B','C','D','E','F','G','H']
window= Tk()
window.geometry('600x500')
window.title("Result Analysis For BTECH")
lb=Label(window,text="WELCOME TO RESULT ANALYSIS",font=("italic",15))
lb.place(x=150,y=80)
lb1=Label(window,text="Enter Source Excel Name:",font=("italic"))
lb1.place(x=100,y=150)
txt=Entry(window,width=30)
txt.place(x=330,y=150)
lb2=Label(window,text="Enter Destination Excel Name:",font=("italic"))
lb2.place(x=100,y=200)
txt1=Entry(window,width=30)
txt1.place(x=330,y=200)
lb3=Label(window,text="Enter Sem No:",font=("italic"))
lb3.place(x=100,y=250)
txt2=Entry(window,width=30)
txt2.place(x=330,y=250)
lb4=Label(window,text="Enter No.Of Subjects:",font=("italic"))
lb4.place(x=100,y=300)
txt3=Entry(window,width=30)
txt3.place(x=330,y=300)
lb5=Label(window,text="Enter Chromedriver Path:",font=("italic"))
lb5.place(x=100,y=350)
txt4=Entry(window,width=30)
txt4.place(x=330,y=350)

lb5=Label(window,text="@copyrights to ECE DEPT SMVEC")
lb5.place(x=400,y=450)
lb6=Label(window,text="Developed By: KISHOREKANNAN S, POOVANAN G")
lb6.place(x=10,y=450)
def clicked5():
    y1=txt.get()
    global z
    z=y1
    y2=txt1.get()
    global kl
    kl=y2
    y3=int(txt2.get())
    global q
    q=y3
    y4=int(txt3.get())
    global sc
    sc=y4
    y5=txt4.get()
    global x
    x=y5
    data = pd.read_excel(z) 
    df = data['Regno'].tolist()
    m="http://exam.pondiuni.edu.in/oresults/result.php?r="
    n="&e="+sem[q-1]
    nf=[]
    pp=[]
    l=[]
    k=[]
    s=[]
    s1=[]
    a0=[]
    a1=[]
    f=[]
    h=[]
    for a in range(len(df)):
        driver = webdriver.Chrome(x)
        driver.get(m+df[a]+n)
        try:
            element1 = driver.find_element_by_id('student_info')
            print(element1.text)
            f=element1.text.split(":")
            if(driver.find_element_by_xpath('//*[@id="results_subject_table"]/tbody')):
                h.append(f[2])
                print("-"*10)
                
                table = driver.find_element_by_xpath('//*[@id="results_subject_table"]/tbody')
                c = 0
                for i in table.find_elements_by_xpath('.//tr'):
                    #print (i.find_element_by_tag_name('td').get_attribute('innerHTML'))
                    if c:
                        for t in i.find_elements_by_tag_name('td'):
                            data = t.get_attribute('innerHTML')
                            if data[:4] == '<div':
                                print(data[data.index('>')+1],end=' ')
                                l.append(data[data.index('>')+1])
                            else:
                                print(data,end=' ')
                                l.append(data)
                            
                    c+=1
                driver.close()
                s.append(h[0])
                for i in range(1,(sc)*7+1,7):
                    k.append(l[i])
                for i in range(6,(sc)*7+1,7):
                    s.append(l[i])
                s1.append(s)
                s=[]
                l=[]
                h=[]
        except:
            pass
            driver.close()

    for j in range(sc+1):
        for i in s1:
                a0.append(i[j])
        a1.append(a0)
        a0=[]
    
    nf.append("No of failures")
    pp.append("Pass  Percentage")
    for i in range(1,len(a1)):
        avg=100.0-((a1[i].count('F')*100.0)/len(df))
        nf.append(a1[i].count('F'))
        pp.append(avg)
    k=k[:sc]
    k.insert(0,"Name")
    s1.insert(0,k)
    s1.insert(len(s1),nf)
    s1.insert(len(s1),pp)
    workbook=xlsxwriter.Workbook(kl+".xlsx")
    worksheet=workbook.add_worksheet()
    for r,rd in enumerate(s1):
        for c,cd in enumerate(rd):
            worksheet.write(r,c,cd)
    workbook.close()
btn5=Button(window,text="Fetch Results",command=clicked5,font=("italic"))
btn5.place(x=250,y=400)
window.mainloop()
