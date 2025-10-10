To run tests, run the following command

1. Clone the repository:

```bash
git clone https://github.com/vinayavr/21_regulation-CO-PO-Automation.git
```

2. Change to project directory:

```bash
cd 21_regulation-CO-PO-Automation
```

3. Setup the project:
\
\
3.1 Create virtual environment
```bash
python -m venv venv 
```
3.2 Activate the virtual environment
```bash
venv\Scripts\activate
```
3.3 Install the dependencies 
```bash
pip install flask pdfplumber pandas openpyxl
```
3.4 Check the dependencies 
```bash
python -m flask --version
```
3.5 Create the flask app 
```bash
set FLASK_APP=COAutomation 
```
4. Running the project:
```bash
flask run
