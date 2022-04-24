import tkinter
from tkinter import*
from tkinter import ttk, filedialog
from tkinter import font as tkFont
import time
from tkinter import *
import os
import shutil
import pandas as pd
import csv
from func_1 import generate_transcripts,check_roll



window=Tk()
window.geometry('1000x500')
window.title('Transcript Generator')



main_path=os.getcwd()
range_area_var=tkinter.StringVar()

#Define Function for opening grade file
def open_file_grade():
    try:
        os.mkdir('input')
    except:
        pass
    file_grade = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if file_grade:
      filepath = os.path.abspath(file_grade.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','grades.csv') )
      except:
          pass
      file_grade.close()


#Function for opening name-roll sheet
def open_file_name_roll():
    try:
        os.mkdir('input')
    except:
        pass
    file_name_roll = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if file_name_roll:
      filepath = os.path.abspath(file_name_roll.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','names-roll.csv') )
      except:
          pass
      file_name_roll.close()


#Function for opening subjects master sheet
def open_file_subject_master():
    try:
        os.mkdir('input')
    except:
        pass
    file_sub_master = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if file_sub_master:
      filepath = os.path.abspath(file_sub_master.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','subjects_master.csv') )
      except:
          pass
      file_sub_master.close()


def open_seal():
    file_seal=filedialog.askopenfile(mode='r', filetypes=[('PNG Files', '*.png')])
    if file_seal:
        filepath=os.path.abspath(file_seal.name)
        try:
            shutil.copy2(filepath, os.path.join(main_path,'input','seal.png'))
        except:
            pass

def open_sign():
    try:
        os.mkdir('input')
    except:
        pass
    file_sign = filedialog.askopenfile(mode='r', filetypes=[('PNG Files', '*.png')])
    if file_sign:
      filepath = os.path.abspath(file_sign.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','signature.png') )
      except:
          pass
      file_sign.close()

def gen_transcript():
    global starting_roll
    global ending_roll
    roll_numbers=range_area_var.get()

    lst = roll_numbers.split("-")
    starting_roll,ending_roll=lst[0],lst[1]
    start_roll=starting_roll.upper()
    end_roll=ending_roll.upper()
    
    folder=os.path.join(os.getcwd(),'transcriptsIITP')
    print(folder)
    #Deleting all files in folder if exists
    for filename in os.listdir(folder):
        os.remove(os.path.join(folder,filename))
    try:
        if os.path.isdir(file_path):
            shutil.rmtree(file_path)
    except:
        pass

    if(check_roll(starting_roll) and check_roll(ending_roll)):
        Label(window, text='INVALID INPUT!', foreground='red').place(x=220,y=160)
    else:
       name_roll=pd.read_csv(os.path.join(main_path,'input','names-roll.csv'))
       subject_master=pd.read_csv(os.path.join(main_path,'input','subjects_master.csv'))
       names_data=open(os.path.join(main_path,'input','grades.csv'),'r')
       names_csv=csv.reader(names_data)
       names_list=[list(record) for record in names_csv][1:]
       missing_nums=generate_transcripts(name_roll, subject_master, names_list,start_roll,end_roll)
       Label(window, text='Created Succesfully!', foreground='green').place(x=250,y=460)
       yc=570
       xc=300
       if(len(missing_nums)>0):
           Label(window, text='Missing Roll Numbers:', foreground='black',font=('Arial',20)).place(x=220,y=520)
           count=1
       for num in missing_nums:
           Label(window, text='{}'.format(num), foreground='red').place(x=xc,y=yc)
           if(count % 5==0):
               xc=xc+100
               yc=540
           count+=1
           yc=yc+30

        

#Heading 
heading_label=Label(window,text='TRANSCRIPT GENERATOR',font=('Arial Black',25)).place(x=550,y=20)
#Browse Button for opening grade sheet
browse_button1=Button(window, text="Browse",command=open_file_grade)
browse_button1.place(x=400,y=120)
lebel1 = Label(window, text="Browse grade sheet",font=('Arial')).place(x=40,y=120)


#Browse Button for opening name-roll sheet
browse_button2=Button(window, text="Browse",command=open_file_name_roll)
browse_button2.place(x=400,y=180)
lebel2 = Label(window, text="Browse names roll number sheet",font=('Arial')).place(x=40,y=180)


#Browse Button for opening subjects master sheet
browse_button3=Button(window, text="Browse",command=open_file_subject_master)
browse_button3.place(x=400,y=240)
lebel3 = Label(window, text="Browse subjects master sheet",font=('Arial')).place(x=40,y=240)


browse_button4=Button(window, text="Browse",command=open_seal)
browse_button4.place(x=400,y=300)
lebel4 = Label(window, text="Browse seal sheet",font=('Arial')).place(x=40,y=300)


browse_button5=Button(window, text="Browse",command=open_sign)
browse_button5.place(x=400,y=360)
lebel5 = Label(window, text="Browse signature",font=('Arial')).place(x=40,y=360)


range_area = Entry(window,textvariable=range_area_var,width = 30).place(x = 400,y = 420)


lebel6 = Label(window, text="range of roll number",font=('Arial')).place(x=40,y=420)

browse_button6=Button(window, text="Submit",command=gen_transcript).place(x=400,y=500)



window.mainloop()
