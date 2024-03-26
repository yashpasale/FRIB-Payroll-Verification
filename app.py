import tkinter as tk
from tkinter import filedialog
from flask import Flask, flash
from werkzeug.utils import secure_filename
import os
import xlwings as xw
import threading
import shutil
from flask import has_request_context

app = Flask(__name__)
app.config['SECRET_KEY'] = 'yashpasale3'

def setup_directories():
    folder_names = ["uploads", "downloads"]
    for folder_name in folder_names:
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
setup_directories()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def run_macro(filepath):
    try:
        wb = xw.Book(filepath)
        selected_sheet = wb.sheets.active
        selected_sheet.range("A1").value = "Yash"
        selected_sheet.range("A2").value = "Pasale"
        modified_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_' + secure_filename(os.path.basename(filepath)))
        wb.save(modified_file_path)
        wb.close()
        if has_request_context():
            flash('Macro executed successfully')
        else:
            print('Macro executed successfully')
        return modified_file_path
    except Exception as e:
        return None

def start_flask_server():
    app.run(debug=False)

flask_thread = threading.Thread(target=start_flask_server)
flask_thread.daemon = True
flask_thread.start()

class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        
        self.title("Excel Macro App")
        self.geometry("500x300")

        self.label = tk.Label(self, text="Upload Excel File")
        self.label.pack(pady=10)

        self.upload_button = tk.Button(self, text="Upload", command=self.upload_file)
        self.upload_button.pack(pady=5)

        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=5)

    def upload_file(self):
        file_path = filedialog.askopenfilename()
        if file_path and os.path.exists(file_path):
            filename = os.path.basename(file_path)
            if allowed_file(filename):
                if has_request_context():
                    flash('File uploaded successfully')
                else:
                    print('File uploaded successfully')
                modified_file_path = run_macro(file_path)
                if modified_file_path:
                    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")], title="Save Modified File As")
                    if save_path:
                        shutil.copy(modified_file_path, save_path)
                        self.status_label.config(text=f"File modified successfully and saved as {save_path}")
                        self.clear_uploads_folder()

                    else:
                        self.status_label.config(text="Save cancelled")
                else:
                    self.status_label.config(text="Failed to modify file")
            else:
                self.status_label.config(text="Invalid file format")
        else:
            self.status_label.config(text="File not found")
    
    def clear_uploads_folder(self):
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"Error deleting file: {e}")


if __name__ == "__main__":
    app_gui = Application()
    app_gui.mainloop()
