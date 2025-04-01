from flask import Blueprint, request, abort, send_from_directory, jsonify, send_file
import os
import pdfplumber
import json
import traceback
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import re


second_bp = Blueprint('second', __name__)


def extract_marks_obtained(pdf_files):
    marks_data = {}
    reg_pattern = re.compile(r"(RA\d{13})")  
    marks_pattern = re.compile(r"(\d{1,3})$")  

    for pdf_path in pdf_files:
        print(f"Processing file: {pdf_path}")

        if not os.path.exists(pdf_path):
            print(f"File not found: {pdf_path}")
            continue

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    raw_text = page.extract_text()
                    if raw_text:
                        for line in raw_text.split("\n"):
                            reg_match = reg_pattern.search(line)
                            marks_match = marks_pattern.search(line)

                            if reg_match and marks_match:
                                reg_no = reg_match.group(1)
                                marks = int(marks_match.group(1))
                                if reg_no in marks_data:
                                    marks_data[reg_no] += marks
                                else:
                                    marks_data[reg_no] = marks
        except Exception as e:
            print(f"Error reading {pdf_path}: {e}")
            traceback.print_exc()

    num_files = len(pdf_files)
    if num_files > 1:
        for reg_no in marks_data:
            marks_data[reg_no] = marks_data[reg_no] / num_files

    return marks_data

def create_excel_sheet(file_path, marks_data, co_splits):
    workbook = Workbook()
    sheet = workbook.active

    
    title_font = Font(bold=True, size=16, color="000000", name="Times New Roman")
    header_font = Font(bold=True, color="000000", name="Times New Roman")
    title_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    header_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    border_style = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    sheet.cell(row=1, column=1, value="CO Allocation").font = title_font
    sheet.cell(row=1, column=1).fill = title_fill
    sheet.cell(row=1, column=1).alignment = header_alignment
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)

    headers = ["S.No", "Register No", "CO1", "CO2", "CO3", "CO4", "CO5", "CO6", "Total"]
    for col_num, header in enumerate(headers, start=1):
        cell = sheet.cell(row=2, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border_style

    for index, (reg_no, total_marks) in enumerate(marks_data.items(), start=1):
        row_index = index + 2
        sheet.cell(row=row_index, column=1, value=index).border = border_style  # S.No
        sheet.cell(row=row_index, column=2, value=reg_no).border = border_style  # Register No

        co_total = 0
        for col, co in enumerate(["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"], start=3):
            co_marks = round(total_marks * (co_splits.get(co, 0) / 100), 2)
            cell = sheet.cell(row=row_index, column=col, value=co_marks)
            cell.border = border_style
            co_total += co_marks

        sheet.cell(row=row_index, column=9, value=co_total).border = border_style  # Total

    workbook.save(file_path)
    print(f"Excel sheet created: {file_path}")

@second_bp.route('/upload2', methods=['POST'])
def upload_files():
    try:
        uploaded_files = request.files.getlist('pdf_files')

        if not uploaded_files:
            return jsonify({'success': False, 'message': 'No files uploaded'}), 400

        try:
            co_splits = {
                "CO1": float(request.form.get("co1", 0)),
                "CO2": float(request.form.get("co2", 0)),
                "CO3": float(request.form.get("co3", 0)),
                "CO4": float(request.form.get("co4", 0)),
                "CO5": float(request.form.get("co5", 0)),
                "CO6": float(request.form.get("co6", 0)),
            }
        except ValueError:
            return jsonify({'success': False, 'message': 'Invalid CO values. Please enter numeric values only.'}), 400

        upload_dir = './uploads'
        results_dir = './static' 
        os.makedirs(upload_dir, exist_ok=True)
        os.makedirs(results_dir, exist_ok=True)

        saved_files = []
        for file in uploaded_files:
            file_path = os.path.join(upload_dir, file.filename)
            file.save(file_path)
            saved_files.append(file_path)

        if not saved_files:
            return jsonify({'success': False, 'message': 'No valid PDFs were saved'}), 400

        marks_data = extract_marks_obtained(saved_files)
        if not marks_data:
            return jsonify({'success': False, 'message': 'No marks data found in PDFs'}), 400

    
        output_file = os.path.join(results_dir, 'co_allocation.xlsx')
        create_excel_sheet(output_file, marks_data, co_splits)

        return jsonify({'success': True, 'message': 'Excel file created successfully', 'download_url': f'/download/co_allocation.xlsx'})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Unexpected error: {str(e)}'}), 500

@second_bp.route("/download/<filename>")
def download_file(filename):
    output_folder = os.path.join(os.getcwd(), "static")  

    file_path = os.path.join(output_folder, filename)
    if os.path.exists(file_path):
        return send_from_directory(output_folder, filename, as_attachment=True)
    else:
        return abort(404, description="File not found")

