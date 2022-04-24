import tkinter
from tkinter import*
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
from tkinter import font as tkFont
import time
from tkinter import *
from pr import individual_marksheet_generater,concise_marksheet_generater,email_send
import os
import shutil

window=Tk()
window.geometry('1000x500')
window.title('Marksheet Generator')
#Declaring variables 
positive_score_var=tkinter.StringVar()
negative_score_var=tkinter.StringVar()


#Label For Heading

heading_label=Label(window, text = 'MARKSHEET GENERATOR',font=("Arial Black", 25)).place(x=550,y=20)
    
# the label for user_password  
positive_score = Label(window, text = "Positive Score",font='Arial').place(x = 250,y = 300)  
    
negative_score = Label(window, text = "Negative score",font='Arial').place(x = 250,y = 350)

positive_score_area = Entry(window, textvariable= positive_score_var ,width = 30).place(x = 400,y = 300)
    
negative_score_area = Entry(window,textvariable=negative_score_var,width = 30).place(x = 400,y = 350)  

global main_path
main_path=os.getcwd()

def uploadFiles():
    pb1 = ttk.Progressbar(window, orient=HORIZONTAL, length=400, mode='determinate')
    pb1.place(x=220,y=250)
    for i in range(5):
        window.update_idletasks()
        pb1['value'] += 20
        time.sleep(1)
    pb1.destroy()
    Label(window, text='File Uploaded Successfully!', foreground='green').place(x=220,y=260)

def ind_gen_marksheet():
    individual_marksheet_generater(positive_score,negative_score)

def con_marksheet():
    concise_marksheet_generater(positive_score,negative_score)

def mail():
    email_send()


def open_file_master():
    try:
        os.mkdir('input')
    except:
        pass
    filename_master = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if filename_master:
      filepath = os.path.abspath(filename_master.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','master_roll.csv') )
      except:
          pass
      filename_master.close()


def open_file_response():
    filename_res = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    try:
        os.mkdir('input')
    except:
        pass
    if filename_res: 
      filepath = os.path.abspath(filename_res.name)
      try:
        shutil.copy2(filepath,os.path.join(main_path,'input','responses.csv') )
      except:
          pass

      filename_res.close()
      


def storing_negative_positive():
    global positive_score
    positive_score=positive_score_var.get()
    global negative_score
    negative_score=negative_score_var.get()

    
# function to change properties of button on hover
def changeOnHover(button, colorOnHover, colorOnLeave):
  
    # adjusting backgroung of the widget
    # background on entering widget
    button.bind("<Enter>", func=lambda e: button.config(background=colorOnHover))
  
    # background color on leving widget
    button.bind("<Leave>", func=lambda e: button.config(background=colorOnLeave))

        
# Create a Button for browsing master_roll file
browse_button1=Button(window, text="Browse",bg='azure',command=open_file_master)
browse_button1.place(x=400,y=120)
changeOnHover(browse_button1, "CadetBlue1", "gray64")


# Add a Label widget
label = Label(window, text="Click the  to browse the master_roll file", font=('Arial')).place(x=40,y=120)

# Create a Button for browsing responses file
browse_button2=Button(window, text="Browse",bg='azure', command=open_file_response)
browse_button2.place(x=400,y=180)
changeOnHover(browse_button2, "CadetBlue1", "gray64")

#Create Upload button
upload_button=Button(window,text='Upload Files',bg='azure',command=uploadFiles)
upload_button.place(x=220,y=220)
changeOnHover(upload_button, "CadetBlue1", "gray64")

# Add a Label widget
labe2 = Label(window, text="Click the  to browse the respones file", font=('Arial')).place(x=40,y=180)

#Button for submitting positive and neagtive marks
submit_button=Button(window,text='Submit',bg='azure',command=storing_negative_positive)
submit_button.place(x=600,y=350)
changeOnHover(submit_button, "CadetBlue1", "gray64")

#Button for concise marksheet

helv36 = tkFont.Font(family='Helvetica', size=16, weight=tkFont.BOLD)
concise_marksheet_button=Button(window, text="Generate Concise marksheet", font=helv36,bg='azure', command=con_marksheet)
concise_marksheet_button.place(x=250,y=460)
changeOnHover(concise_marksheet_button, "CadetBlue1", "gray64")

#button for Individual marksheet
individual_marksheet_button=Button(window, text="Generate Individual marksheet", font=helv36,bg='azure',command=ind_gen_marksheet) 
individual_marksheet_button.place(x=250,y=550)
changeOnHover(individual_marksheet_button, "CadetBlue1", "gray64")

#Button for Sending mail
send_mail_button=Button(window, text="Send Mail", font=helv36,bg='azure', command=mail)
send_mail_button.place(x=600,y=650)
changeOnHover(send_mail_button, "CadetBlue1", "gray64")


# for the window to display
window.mainloop()
