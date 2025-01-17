from flask import Flask, render_template
import os

import COAutomation

app = Flask(__name__)

@app.route('/')
def index():
    working_path=os.getcwd()
    # pdf_paths = [working_path + "/input/MTech-DSA-CT1-SetA.pdf", working_path + "/input/MTech-DSA-CT2-SetA.docx-2.pdf"] 
    pdf_paths = [working_path + "/input/DSA-CT1-SetA.pdf", working_path + "/input/DSA-CT3-SetA.pdf",working_path + "/input/DSA-CT3-SetA.pdf"]  # Input PDF file path
    COAutomation.generate_excel(pdf_paths)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
