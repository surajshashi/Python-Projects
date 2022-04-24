import func
from fpdf import FPDF
import csv,os,re
import pandas as pd
from func import generate_data
from werkzeug.utils import secure_filename
from datetime import date

input_path = "./input"
sample_output_path = "./transcriptsIITP"

def handle_file_save(FileObject,req_resp_name):
    input= os.path.join(os.getcwd(),"input")
    os.makedirs(input,exist_ok = True)
    if FileObject.filename :
        file_name = FileObject.filename
        filename = secure_filename(file_name)
        FileObject.save(os.path.join(input_path,req_resp_name))
        return True
    elif os.path.exists(f"./input/{req_resp_name}"): return True
    else : return False 
     

def check_roll(rollno):
    pattern = re.compile(r'\d\d\d\d\w\w\d\d') 
    check = re.search(pattern,rollno)
    if check == None:
        return True
    return False

def check_files():
    c1 = os.path.exists('./input/grades.csv')
    c2 = os.path.exists('./input/names-roll.csv')
    c3 = os.path.exists('./input/subjects_master.csv')
    return c1 and c2 and c3

def generate_transcripts(name_roll,subject_master,names_list,start_roll,end_roll):
    credit_dict = {'AA':10,'AB':9,'BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'F':0,'I':0,
                'AA*':10,'AB*':9,'BB*':8,'BC*':7,'CC*':6,'CD*':5,'DD*':4,'F*':0,'I*':0}
    courses_dict = {"CS":"Computer Science and Technology","EE":"Electrical Engineering","ME":"Mechanical Engineering","CE":"Civil and Environmental Engineering","CB":"Chemical Engineering","MM":"Metallurgical and Materials Engineering"}
    table_dict= generate_data(name_roll,subject_master,names_list)
    missing_nums = []
    if start_roll:
        name_roll,missing_nums= func.generate_rollno_list(name_roll,start_roll,end_roll)
    for index,row in name_roll.iterrows():
        pdf = FPDF("L" , "mm" ,"A3")
        pdf.add_page()
        roll,name,cpi= row["Roll"],row["Name"],0
        func.generate_header_layout(pdf,roll,name,2000+int(roll[0:2]),'Btech',courses_dict[roll[4:6]])
        pdf.set_font("Times", size=8)
        col_width_list = [15,70,13,10,10]
        coll_width = sum(col_width_list)
        stx,sty,mth,count,line_height= 19.8,60,0,1,4
        while count<=8:
            if (roll,str(count)) not in table_dict:
                break
            data = table_dict[roll,str(count)]
            credits = sum([item[3]*credit_dict[item[4].strip()] for item in data[1:]])
            total_credits = sum([item[3] for item in data[1:]])
            spi = credits/total_credits
            cpi+=spi
            prestx = stx
            pdf.set_xy(stx,sty)
            pdf.set_font("Times",'B',size=8)
            pdf.cell(10,7,f"Semester {count}")
            pdf.set_xy(stx,sty+7)
            sum1 = sty+7
            for row in data:   
                for ind,datum in enumerate(row):
                    pdf.multi_cell(col_width_list[ind], line_height, str(datum), border=1,align="C", ln=3, max_line_height=pdf.font_size)
                pdf.set_font("Times",'',size=8)
                sum1+=line_height
                pdf.set_xy(stx,sum1)
            stx+=coll_width+10
            pdf.rect(prestx,sum1+2,100,7)
            func.generate_cpi_credits(prestx,sum1+2,100,pdf,spi,cpi/count,total_credits)
            mth = max(mth,sum1+10)
            if count%3==0:
                stx=19.8
                pdf.rect(10,mth,390,0)
                sty=mth+2
            count+=1

        func.generate_footer_layout(mth,pdf)
        try:
            os.mkdir('transcriptsIITP')
        except:
            pass
        pdf.output('./transcriptsIITP/{}.pdf'.format(roll))
    return missing_nums


