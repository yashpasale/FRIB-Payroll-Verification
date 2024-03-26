from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename
import os
import xlwings as xw

app = Flask(__name__)
app.config['SECRET_KEY'] = 'yashpasale3'

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def run_macro(filepath):
    try:
        # Open the Excel file
        wb = xw.Book(filepath)
        
        # Get the active sheet
        selected_sheet = wb.sheets.active
        
        # Run the macro
        selected_sheet.range("A1").value = "Yash"
        selected_sheet.range("A2").value = "Pasale"
        # Save the changes
        modified_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_' + secure_filename(os.path.basename(filepath)))
        wb.save(modified_file_path)
        
        # Close the workbook
        wb.close()
        
        flash('macro executed successfully')
        return modified_file_path
    except Exception as e:
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        return redirect(request.url)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        flash('file uploaded successfully')
        modified_file_path = run_macro(file_path)
        if modified_file_path:
            return redirect(url_for('download_file', filename=os.path.basename(modified_file_path)))
        else:
            return redirect(request.url)
    else:
        return redirect(request.url)
    
@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            flash('File not found')
            return redirect(url_for('index'))

        # Send the file as an attachment
        response = send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

        # Delete all files in the upload folder
        files_to_delete = os.listdir(app.config['UPLOAD_FOLDER'])
        for file_to_delete in files_to_delete:
            file_path_to_delete = os.path.join(app.config['UPLOAD_FOLDER'], file_to_delete)
            if os.path.isfile(file_path_to_delete):
                os.remove(file_path_to_delete)

        return response
        
    except Exception as e:
        print("Error:", e)
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
