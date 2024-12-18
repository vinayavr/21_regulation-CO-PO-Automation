from flask import Flask, render_template, make_response,request, jsonify, send_file
from openpyxl import Workbook, styles,load_workbook
import pandas as pd
import os
from openpyxl.styles import Font,Border
from openpyxl.utils import get_column_letter
import pdfplumber
import PyPDF2
import json
import re
import pytesseract
from PIL import Image
from fpdf import FPDF
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/')
def index():
    working_path=os.getcwd()
    pdf_paths = [working_path + "/input/DSA-CT3-SetA.pdf",working_path + "/input/DSA-CT3-SetA.pdf",working_path + "/input/DSA-CT3-SetA.pdf"]  # Input PDF file path
    generate_excel(pdf_paths)
    return render_template('index.html')

def extract_text_from_pdf(pdf_path):
    text_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                text_lines.extend(lines)
    return text_lines

def extract_info_from_text(text_lines):
    info = {
        "Program Section:": "",
        "Subject Code & Title:": "",
        "Test Name:": ""
    }
    for line in text_lines:
        for key in info.keys():
            if line.startswith(key):
                info[key] = line
    return info

def extract_questions_from_text(text_lines,questions):
    q_dict = {}
    for line in text_lines:
        for Q in questions:
            if Q in line:
                parts = line.split()
                Q_index = parts.index(Q)
                if Q_index + 1 < len(parts):
                    q_dict[Q] = parts[Q_index + 1]
    return q_dict

# Function to process PDFs in a folder
def process_pdfs_in_folder(file_order, folder_path, reg_numbers):
    results = {}
    info = {
        "Program Section:": "",
        "Subject Code & Title:": "",
        "Test Name:": ""
    }
    for filename in file_order:
        pdf_path = os.path.join(folder_path, filename)
        text_lines = extract_text_from_pdf(pdf_path)
        marks = extract_questions_from_text(text_lines, reg_numbers)
        for reg_number, marks_value in marks.items():
            if reg_number not in results:
                results[reg_number] = {}
            results[reg_number][filename] = marks_value
        extracted_info = extract_info_from_text(text_lines)
        for key in info:
            if extracted_info[key]:
                info[key] = extracted_info[key]
    return results, info

def extract_question_numbers_from_pdf(pdf_path, question_numbers, marks):
    """
    Extracts question numbers (Q.no) from a question paper PDF.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        list: A list of extracted question numbers as integers.
    """

    # Define a regex pattern to match question numbers (e.g., "1", "2", etc.)
    flag=0

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            # Extract all text from the page
            text = page.extract_text()

            if text: 
                text = text.replace('&', '"&"')
                 # Ensure the page has text
                # Split the text into lines
                lines = text.split("\n")

                for line in lines:
                    if flag==0 and line.find("Part")!=-1:
                        flag=1
                        continue
                    if flag==1:
                        match = re.search(r".\d \d \d \d$",line.strip()) 
                        question_no= re.match(r"^(\d{2})|\d",line.strip()) 
                        if question_no and match:
                            # Add the matched number to the list
                            try:
                                question_numbers.append(int(question_no.group()))
                                marks.append(int(match.group().strip().split()[0]))
                            except ValueError:
                                pass  # Skip if the match isn't a valid integer

def generate_question_number(question_numbers,length):
    next=0
    flag=0
    qnum=[]
    for i in range(0, length):
        if i!=length-1:
            next=question_numbers[i+1]
        else:
            next=0
        if question_numbers[i]==next:
            qnum.append("Q"+str(question_numbers[i])+".A")
            flag=1
        elif flag==1:
            qnum.append("Q"+str(question_numbers[i])+".B")
            flag=0
        else:
            qnum.append("Q"+str(question_numbers[i]))
    return qnum
    

def generate_excel(pdf_paths):

    question_numbers1=[]
    question_numbers2=[]
    question_numbers3=[]
    marks1=[]
    marks2=[] 
    marks3=[]

    extract_question_numbers_from_pdf(pdf_paths[0], question_numbers1, marks1)
    extract_question_numbers_from_pdf(pdf_paths[1], question_numbers2, marks2)
    extract_question_numbers_from_pdf(pdf_paths[2], question_numbers3, marks3)

    # Create a new workbook
    workbook = Workbook()
    worksheet = workbook.active

# Define colors
    light_green_fill = styles.PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    orange_fill = styles.PatternFill(start_color="FFC300", end_color="FFC300", fill_type="solid")
    yellow_fill = styles.PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    grey_fill = styles.PatternFill(start_color="c0c0c0", end_color="c0c0c0", fill_type="solid")

# Define a bold border  
    bold_border = styles.Border(left=styles.Side(border_style='thin', color='000000'),
                    right=styles.Side(border_style='thin', color='000000'),
                    top=styles.Side(border_style='thin', color='000000'),
                    bottom=styles.Side(border_style='thin', color='000000'))

    co1=len(question_numbers1)
    co2=len(question_numbers2)
    co3=len(question_numbers3)

    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    worksheet.cell(row=1, column=1).value = "CLAT->"
    worksheet.cell(row=1, column=1).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=1, column=1).fill = light_green_fill

    worksheet.merge_cells(start_row=1, start_column=4, end_row=1, end_column=3+co1)
    worksheet.cell(row=1, column=4).value = "FT-I"
    worksheet.cell(row=1, column=4).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=1, column=4).fill = light_green_fill

    worksheet.merge_cells(start_row=1, start_column=4+co1, end_row=1, end_column=3+co1+co2)
    worksheet.cell(row=1, column=4+co1).value='FT-II , FT-IV'
    worksheet.cell(row=1, column=4+co1).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=1, column=4+co1).fill=light_green_fill

    worksheet.merge_cells(start_row=1, start_column=4+co1+co2, end_row=1, end_column=3+co1+co2+co3)
    worksheet.cell(row=1, column=4+co1+co2).value='FT-III'
    worksheet.cell(row=1, column=4+co1+co2).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=1, column=4+co1+co2).fill=light_green_fill

    worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    worksheet.cell(row=2, column=1).value = "CO ->"
    worksheet.cell(row=2, column=1).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=2, column=1).fill =orange_fill
    
    worksheet.merge_cells(start_row=2, start_column=4, end_row=2, end_column=3+co1)
    worksheet.cell(row=2, column=4).value = "CO1"
    worksheet.cell(row=2, column=4).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=2, column=4).fill = orange_fill
    
    worksheet.merge_cells(start_row=2, start_column=4+co1, end_row=2, end_column=3+co1+co2)
    worksheet.cell(row=2, column=4+co1).value='CO2'
    worksheet.cell(row=2, column=4+co1).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=2, column=4+co1).fill=orange_fill

    worksheet.merge_cells(start_row=2, start_column=4+co1+co2, end_row=2, end_column=3+co1+co2+co3)
    worksheet.cell(row=2, column=4+co1+co2).value='CO3'
    worksheet.cell(row=2, column=4+co1+co2).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=2, column=4+co1+co2).fill=orange_fill

    worksheet.merge_cells(start_row=3, start_column=4, end_row=3, end_column=3+co1)
    worksheet.cell(row=3, column=4).value = 'THEORY \n (for either/or Q, \n award marks for the attempted students only)'
    worksheet.cell(row=3, column=4).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.merge_cells(start_row=3, start_column=4+co1, end_row=3, end_column=3+co1+co2)
    worksheet.cell(row=3, column=4+co1).value='THEORY \n (for either/or Q, \n award marks for the attempted students only)'
    worksheet.cell(row=3, column=4+co1).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.merge_cells(start_row=3, start_column=4+co1+co2, end_row=3, end_column=3+co1+co2+co3)
    worksheet.cell(row=3, column=4+co1+co2).value='THEORY \n (for either/or Q, \n award marks for the attempted students only)'
    worksheet.cell(row=3, column=4+co1+co2).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.merge_cells(start_row=4, start_column=1, end_row=4, end_column=3)
    worksheet.cell(row=4, column=1).value = "MAX. MARKS (If not applicable, leave BLANK)->"
    worksheet.cell(row=4, column=1).alignment = styles.Alignment(horizontal='right', vertical='center')
    worksheet.cell(row=4, column=1).fill = yellow_fill

    for i in range(0, co1):
        worksheet.column_dimensions[get_column_letter(4+i)].width=8
        worksheet.cell(row=4, column=4+i).value = marks1[i]
        worksheet.cell(row=4, column=4+i).alignment = styles.Alignment(horizontal='center', vertical='center')

    for i in range(0, co2):
        worksheet.column_dimensions[get_column_letter(4+co1+i)].width=8
        worksheet.cell(row=4, column=4+co1+i).value = marks2[i]
        worksheet.cell(row=4, column=4+co1+i).alignment = styles.Alignment(horizontal='center', vertical='center')
    
    for i in range(0, co3):
        worksheet.column_dimensions[get_column_letter(4+co1+co2+i)].width=8
        worksheet.cell(row=4, column=4+co1+co2+i).value = marks3[i]
        worksheet.cell(row=4, column=4+co1+co2+i).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.merge_cells(start_row=5, start_column=4, end_row=5, end_column=3+co1)
    worksheet.cell(row=5, column=4).value = 'Question numbers mapping'
    worksheet.cell(row=5, column=4).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=5, column=4).fill = grey_fill

    worksheet.merge_cells(start_row=5, start_column=4+co1, end_row=5, end_column=3+co1+co2)
    worksheet.cell(row=5, column=4+co1).value = 'Question numbers mapping'
    worksheet.cell(row=5, column=4+co1).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=5, column=4+co1).fill = grey_fill

    worksheet.merge_cells(start_row=5, start_column=4+co1+co2, end_row=5, end_column=3+co1+co2+co3)
    worksheet.cell(row=5, column=4+co1+co2).value = 'Question numbers mapping'
    worksheet.cell(row=5, column=4+co1+co2).alignment = styles.Alignment(horizontal='center', vertical='center')
    worksheet.cell(row=5, column=4+co1+co2).fill = grey_fill

    worksheet.cell(row=6,column=1).value="Sl.No"
    worksheet.column_dimensions[get_column_letter(1)].width=6
    worksheet.cell(row=6, column=1).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.cell(row=6,column=2).value="Register Number"
    worksheet.column_dimensions[get_column_letter(2)].width=20
    worksheet.cell(row=6, column=2).alignment = styles.Alignment(horizontal='center', vertical='center')

    worksheet.cell(row=6,column=3).value="student Name"
    worksheet.column_dimensions[get_column_letter(3)].width=30
    worksheet.cell(row=6, column=3).alignment = styles.Alignment(horizontal='center', vertical='center')

    qnum=generate_question_number(question_numbers1,co1)
    for i in range(0, co1):
        worksheet.cell(row=6, column=4+i).value = qnum[i]
        worksheet.cell(row=6, column=4+i).alignment = styles.Alignment(horizontal='center', vertical='center')
    
    qnum=generate_question_number(question_numbers2,co2)        
    for i in range(0, co2):
        worksheet.cell(row=6, column=4+co1+i).value = qnum[i]
        worksheet.cell(row=6, column=4+co1+i).alignment = styles.Alignment(horizontal='center', vertical='center')
    
    qnum=generate_question_number(question_numbers3,co3)
    for i in range(0, co3):
        worksheet.cell(row=6, column=4+co1+co2+i).value =qnum[i]
        worksheet.cell(row=6, column=4+co1+co2+i).alignment = styles.Alignment(horizontal='center', vertical='center')
    
    times_new_roman_font = Font(name="Times New Roman", size=10, bold=True)
    for row_num, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row), 1):
            for cell in row:
                cell.font = times_new_roman_font
                cell.border=bold_border
                if row_num == 4:
                    cell.fill = yellow_fill
            

# Adjust the path as needed
    file_path = os.getcwd()+"/output/result.xlsx" 
    # Save the workbook to the specified path
    workbook.save(file_path)

    # Prepare a success message (optional)
    message = "Excel file saved successfully to: " +  "result.xlsx"

    # Return the message (no download functionality)
    return render_template("index.html", message=message)

if __name__ == '__main__':
    app.run(debug=True)