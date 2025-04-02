import traceback
import os
import sys
import json
import logging
from typing import Dict, List, Optional

import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re

from flask import Blueprint, request, abort, send_from_directory,render_template, jsonify, send_file, current_app, Flask

second_bp = Blueprint('second', __name__)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class TLPMarkConverter:
    def __init__(self, config: Dict = None):
        self.config = config or {
            'upload_dir': './uploads',
            'results_dir': './static',
            'max_file_size': 50 * 1024 * 1024,
            'allowed_extensions': {'.pdf'}
        }
        
        os.makedirs(self.config['upload_dir'], exist_ok=True)
        os.makedirs(self.config['results_dir'], exist_ok=True)

    def validate_file(self, filename: str) -> bool:
        return os.path.splitext(filename)[1].lower() in self.config['allowed_extensions']

    def extract_marks_from_tlp(self, pdf_files: List[str]) -> Dict[str, Dict]:
        """
        Extract marks from multiple PDF files, aggregating by register number
        
        :param pdf_files: List of PDF file paths
        :return: Dictionary containing marks data and file processing statistics
        """
        marks_data = {}
        stats = {
            'total_files': len(pdf_files),
            'processed_files': 0,
            'failed_files': 0,
            'total_entries': 0,
            'file_stats': {}
        }
        
        # Regular expression for conducted max value
        conducted_max_pattern = re.compile(r"Conducted Max\.?\s+(\d+\.?\d*)")
        
        # Regular expression for register numbers and marks (handles A and 0A)
        reg_pattern = re.compile(r"(RA\d{13})\s+((?:\d+\.?\d*)|(?:[A0]A))")
        
        for pdf_path in pdf_files:
            file_name = os.path.basename(pdf_path)
            entries_found = 0
            conducted_max = None
            
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    file_text = ""
                    for page in pdf.pages:
                        # Extract full text and append
                        page_text = page.extract_text() or ""
                        file_text += page_text
                    
                    # Find conducted max value first
                    max_match = conducted_max_pattern.search(file_text)
                    if max_match:
                        conducted_max = float(max_match.group(1))
                        logger.info(f"Found Conducted Max value: {conducted_max} in {file_name}")
                    else:
                        logger.warning(f"Conducted Max value not found in {file_name}")
                    
                    # Find all register number and mark matches
                    matches = reg_pattern.findall(file_text)
                    
                    for reg_no, marks in matches:
                        # Process marks value
                        if marks == 'A' or marks == '0A':
                            current_marks = 0
                        else:
                            current_marks = float(marks)
                        
                        # Aggregate marks for each register number
                        if reg_no in marks_data:
                            marks_data[reg_no] += current_marks
                        else:
                            marks_data[reg_no] = current_marks
                        
                        entries_found += 1
                
                stats['processed_files'] += 1
                stats['total_entries'] += entries_found
                stats['file_stats'][file_name] = {
                    'status': 'success',
                    'entries_found': entries_found,
                    'conducted_max': conducted_max
                }
                logger.info(f"Successfully processed {file_name}: found {entries_found} entries")
                
            except Exception as e:
                stats['failed_files'] += 1
                stats['file_stats'][file_name] = {
                    'status': 'failed',
                    'error': str(e)
                }
                logger.error(f"Error processing PDF {pdf_path}: {e}", exc_info=True)
        
        logger.info(f"Total unique entries found across all files: {len(marks_data)}")
        return {
            'marks_data': marks_data,
            'stats': stats
        }

    def create_excel_sheet(
        self, 
        file_path: str, 
        marks_data: Dict[str, float], 
        co_splits: Dict[str, int],
        processing_stats: Dict = None
    ) -> None:
        """
        Generate styled Excel spreadsheet with CO mark distribution
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "CO Mark Distribution"
        
        # Create a summary sheet for file processing stats
        summary_sheet = workbook.create_sheet(title="Processing Summary")
        
        styles = {
            'title_font': Font(bold=True, size=16, name="Arial"),
            'header_font': Font(bold=True, name="Arial"),
            'title_fill': PatternFill(start_color="92D050", fill_type="solid"),
            'header_fill': PatternFill(start_color="F0F0F0", fill_type="solid"),
            'border': Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )
        }
        
        # Main Sheet - Title and headers
        sheet.merge_cells('A1:I1')
        title_cell = sheet.cell(row=1, column=1, value="CO Mark Distribution")
        title_cell.font = styles['title_font']
        title_cell.fill = styles['title_fill']
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        headers = ["S.No", "Register No", "Total Marks", "CO1", "CO2", "CO3", "CO4", "CO5", "CO6"]
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=2, column=col, value=header)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = Alignment(horizontal='center')
            cell.border = styles['border']

        sheet.column_dimensions[get_column_letter(1)].width=6
        sheet.column_dimensions[get_column_letter(2)].width=20

        # Sort marks data by register number
        sorted_marks = sorted(marks_data.items(), key=lambda x: x[0])
        
        # Populate data rows
        for index, (reg_no, total_marks) in enumerate(sorted_marks, 1):
            row = index + 2
            sheet.cell(row=row, column=1, value=index).border = styles['border']
            sheet.cell(row=row, column=2, value=reg_no).border = styles['border']
            sheet.cell(row=row, column=3, value=total_marks).border = styles['border']
            
            # Calculate CO marks based on total marks and percentages
            co_total = sum(co_splits.get(f'CO{i}', i) for i in range(1, 7))
            
            for col, co in enumerate(["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"], 4):
                # Calculate proportional marks for this CO
                co_value = co_splits.get(co, 0)
                if co_total > 0:
                    co_marks = (co_value / co_total) * total_marks
                else:
                    co_marks = 0
                
                cell = sheet.cell(row=row, column=col, value=round(co_marks, 2))
                cell.border = styles['border']
                cell.alignment = Alignment(horizontal='center')
        
        # Add total CO marks summary at the bottom
        total_row = len(sorted_marks) + 3
        sheet.cell(row=total_row, column=1, value="Total CO Marks:").font = styles['header_font']
        for col, co in enumerate(["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"], 4):
            sheet.cell(row=total_row, column=col, value=co_splits.get(co, 0)).font = styles['header_font']
        
        # Auto-adjust column widths
        for col in range(1, 10):
            sheet.column_dimensions[get_column_letter(col)].auto_size = True
        
        # Summary Sheet - Processing Statistics
        if processing_stats:
            # Title
            summary_sheet.merge_cells('A1:E1')
            summary_title = summary_sheet.cell(row=1, column=1, value="File Processing Summary")
            summary_title.font = styles['title_font']
            summary_title.fill = styles['title_fill']
            summary_title.alignment = Alignment(horizontal='center', vertical='center')
            
            # Overall stats
            summary_sheet.cell(row=2, column=1, value="Total Files:").font = styles['header_font']
            summary_sheet.cell(row=2, column=2, value=processing_stats.get('total_files', 0))
            
            summary_sheet.cell(row=3, column=1, value="Successfully Processed:").font = styles['header_font']
            summary_sheet.cell(row=3, column=2, value=processing_stats.get('processed_files', 0))
            
            summary_sheet.cell(row=4, column=1, value="Failed Files:").font = styles['header_font']
            summary_sheet.cell(row=4, column=2, value=processing_stats.get('failed_files', 0))
            
            summary_sheet.cell(row=5, column=1, value="Total Unique Entries:").font = styles['header_font']
            summary_sheet.cell(row=5, column=2, value=processing_stats.get('total_entries', 0))
            
            # File details table headers
            headers = ["S.No", "Filename", "Status", "Entries Found", "Error (if any)"]
            for col, header in enumerate(headers, 1):
                cell = summary_sheet.cell(row=7, column=col, value=header)
                cell.font = styles['header_font']
                cell.fill = styles['header_fill']
                cell.alignment = Alignment(horizontal='center')
                cell.border = styles['border']
            
            # File details
            file_stats = processing_stats.get('file_stats', {})
            row = 8
            for idx, (filename, stats) in enumerate(file_stats.items(), 1):
                summary_sheet.cell(row=row, column=1, value=idx).border = styles['border']
                summary_sheet.cell(row=row, column=2, value=filename).border = styles['border']
                summary_sheet.cell(row=row, column=3, value=stats.get('status', 'unknown')).border = styles['border']
                summary_sheet.cell(row=row, column=4, value=stats.get('entries_found', 0)).border = styles['border']
                summary_sheet.cell(row=row, column=5, value=stats.get('error', '')).border = styles['border']
                row += 1
            
            # Auto-adjust column widths
            for col in range(1, 6):
                summary_sheet.column_dimensions[get_column_letter(col)].auto_size = True
        
        workbook.save(file_path)
        logger.info(f"Excel sheet created: {file_path}")

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
        co_splits={}
        for i in range(1,7):
            co_value = request.form.get('co'+str(i),'')
            if co_value=='':
                co_splits['CO'+str(i)]=0
            else:
                co_splits['CO'+str(i)]=float(co_value)
        

        if not uploaded_files or uploaded_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No files uploaded'}), 400

        converter = TLPMarkConverter()
        saved_files = []
        invalid_files = []
        
        for file in uploaded_files:
            if file and file.filename:
                if converter.validate_file(file.filename):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(converter.config['upload_dir'], filename)
                    file.save(file_path)
                    saved_files.append(file_path)
                else:
                    invalid_files.append(file.filename)

        if not saved_files:
            return jsonify({
                'success': False, 
                'message': f'No valid PDFs were saved. Invalid files: {", ".join(invalid_files)}'
            }), 400
        

        # Extract marks from all uploaded files
        extraction_result = converter.extract_marks_from_tlp(saved_files)
        marks_data = extraction_result['marks_data']
        stats = extraction_result['stats']
        
        if not marks_data:
            return jsonify({
                'success': False, 
                'message': 'No marks data found in PDFs. Please check the file format.'
            }), 400
        
        total_conducted_max = 0.0
    
        # Access the file_stats dictionary in the returned data
        file_stats = extraction_result.get('stats', {}).get('file_stats', {})
        
        # Iterate through each file's stats and sum up the conducted_max values
        for file_name, stats in file_stats.items():
            if stats.get('status') == 'success' and 'conducted_max' in stats:
                file_conducted_max = stats['conducted_max']
                if file_conducted_max is not None:  # Check if the value was found
                    total_conducted_max += file_conducted_max
                    logger.info(f"Added {file_conducted_max} from {file_name} to total")
                else:
                    logger.warning(f"No Conducted Max found for {file_name}")
        
        logger.info(f"Total Conducted Max across all files: {total_conducted_max}")

        co_total = sum(co_splits.get(f'CO{i}', i) for i in range(1, 7))

        if co_total!=total_conducted_max:
            return jsonify({
                'success': False, 
                'message': 'CO split is not proper.please enter correct splitup'
            }), 400
        
        output_file = os.path.join(converter.config['results_dir'], 'co_allocation.xlsx')
        converter.create_excel_sheet(output_file, marks_data, co_splits, stats)
        
        # Prepare response message
        success_message = (
            f"Excel file created successfully with {len(marks_data)} unique entries. "
            f"Processed {stats['processed_files']} out of {stats['total_files']} files."
        )
        
        if stats['failed_files'] > 0:
            success_message += f" {stats['failed_files']} files could not be processed. Check the Processing Summary sheet for details."
        
        if invalid_files:
            success_message += f" Ignored {len(invalid_files)} non-PDF files."
        
        return jsonify({
            'success': True, 
            'message': success_message,
            'processed_files': stats['processed_files'],
            'failed_files': stats['failed_files'],
            'total_entries': len(marks_data),
            'download_url': '/download/co_allocation.xlsx'
        })
    
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Unexpected error: {str(e)}'}), 500

@second_bp.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(current_app.config['results_dir'], filename)
    return send_file(file_path, as_attachment=True) if os.path.exists(file_path) else ("File not found", 404)

