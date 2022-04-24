from fpdf import FPDF
import csv,os,re
import pandas as pd
from datetime import date
import flask
from flask import Flask , flash , request , send_file , render_template ,  Markup
import os
import shutil
import csv
import datetime
from datetime import datetime



def generate_header_layout(pdf,*stud_info):
    pdf.set_font("Times",'B',size=10)
    pdf.image('./logo.png' ,15,11,30,25)
    pdf.image('./head_text.png',50,11,300,25)
    pdf.image('./logo.png' ,365,11,30,25)

    pdf.rect(10,10,390,277)
    pdf.rect(10,40,390,0)
    pdf.rect(52,11,0,29)
    pdf.rect(360,11,0,29)
    pdf.rect(95,43,240,13)
    
    lst,x,y= ["INTERIM TRANSCRIPT","TRANSCRIPT","INTERIM TRANSCRIPT"],10,36
    for item in lst:
        pdf.set_xy(x,36)
        if item == "TRANSCRIPT": pdf.set_font("Times",'B',size=15)
        else:pdf.set_font("Times",'B',size=10)
        pdf.cell(25,5,item)
        x+=175
    x=0
    for index,item in enumerate(["Roll No:","Name:","Year of Admission:","Programme:","Course:"]):
        x+=95
        k=((index)//3) * 6
        pdf.set_xy(x,43+k)
        pdf.set_font("Times",'B',size=10)
        pdf.cell(10,7,item)
        pdf.set_x(x+2*len(item))
        pdf.set_font("Times",'',size=10)
        pdf.cell(10,7,str(stud_info[index]))
        x%=285
    
    return 


def generate_cpi_credits(x,y,w,pdf,spi,cpi,total_credits):
    pdf.set_font("Times",'B',size=8)
    summ_list = [f"Credits taken: {total_credits}",f"Credits cleared: {total_credits}",f"SPI: {round(spi,2)}",f" CPI:{round(cpi,2)}"]
    for item in summ_list:
        pdf.set_xy(x,y)
        pdf.cell(10,7,item) 
        x+=(w/4)
    return


def generate_footer_layout(mth,pdf):
    pdf.rect(10,mth,390,0)
    pdf.set_xy(19.8,mth+(287-mth)/2)
    pdf.cell(15,7,"Date of Issue")
    pdf.rect(pdf.get_x()+5,pdf.get_y()+5,30,0)
    pdf.set_xy(pdf.get_x()+5,pdf.get_y())
    pdf.cell(15,7,str(date.today()))
    if os.path.exists('./input/seal.png'):
        pdf.image('./input/seal.png' ,pdf.get_x()+150,mth+10,50,50)
    pdf.rect(350,pdf.get_y(),30,0)
    pdf.set_xy(350,pdf.get_y())
    if os.path.exists('./input/sign.png'):
        pdf.image('./input/signature.png' ,350,pdf.get_y()-12,30,25)
    pdf.cell(15,7,"Assistant Registrar(Academic)")
    return

def generate_data(name_roll,subject_master,names_list):   
    name_roll_dict = {}
    for i in range(len(name_roll)): name_roll_dict[name_roll.at[i,"Roll"]] = name_roll.at[i,"Name"]
    subject_master_dict = {}
    for i in range(len(subject_master)): 
        subject_master_dict[subject_master.at[i,"subno"]] = [subject_master.at[i,"subno"],subject_master.at[i,"subname"],subject_master.at[i,"ltp"],subject_master.at[i,"crd"]]   
    table_dict = {}
    for row in names_list:
        rollno,semno,subcode,credit,grade,Sub_Type = row
        st_list = subject_master_dict[f"{subcode}"].copy()
        grade = str(grade)
        st_list.append(grade)
        if (rollno,semno) not in table_dict:
            table_dict[rollno,semno] = [["Sub Code","Subject Name","L-T-P","CRD","GRD"]]
        table_dict[rollno,semno].append(st_list)
    return table_dict

def generate_rollno_list(name_roll,start_roll,end_roll):
    name_roll_dict,missing_roll,existing_nums={},[],[]
    
    for i in range(len(name_roll)): name_roll_dict[name_roll.at[i,"Roll"]] = name_roll.at[i,"Name"]
    starting_roll,ending_roll = int(start_roll[6:]),int(end_roll[6:])
    if starting_roll>ending_roll : starting_roll,ending_roll = ending_roll,starting_roll
    st = start_roll[:6]
    
    for i in range(starting_roll,ending_roll+1):
        if len(str(i))==1: num = "0"+str(i)
        else: num = str(i)
        rollno = st+num
        if rollno not in name_roll_dict:
            missing_roll.append(rollno)
        else :
            existing_nums.append([rollno,name_roll_dict[rollno]]) 
    return pd.DataFrame(existing_nums,columns = ["Roll","Name"]),missing_roll

