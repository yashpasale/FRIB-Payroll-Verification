import tkinter as tk
from tkinter import filedialog, ttk
import os
import shutil
import threading
import pandas as pd
import xlwings as xw
import re
import openpyxl

app = tk.Tk()
app.title("FRIB Payroll Validation")
app.geometry("500x300")

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls', 'txt'}

def get_total_rows(filepath):
    try:
        wb = xw.Book(filepath)
        lastrow = wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
        return lastrow
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def align_data(text_content):
    lines = text_content.split('\n')
    prefixes = [re.match(r'^\d+', line).group() for line in lines if re.match(r'^\d+', line)]
    max_prefix_len = max(len(prefix) for prefix in prefixes)
    padded_prefixes = [prefix.zfill(max_prefix_len) for prefix in prefixes]
    aligned_lines = [f"{padded_prefix}: {line[len(prefix):]}" for prefix, padded_prefix, line in
                     zip(prefixes, padded_prefixes, lines)]
    return '\n'.join(aligned_lines)


def txt_to_excel_file(file_path):
    try:
        if os.path.exists(file_path):
            with open(file_path, 'r') as text_file:
                text_content = text_file.read()

            aligned_content = align_data(text_content)
            data = [re.split(r'\s+', line.strip()) for line in text_content.split('\n')]
            df = pd.DataFrame(data)
            excel_save_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                           filetypes=[("Excel Files", "*.xlsx")],
                                                           title="Save Excel File As")
            if excel_save_path:
                df.to_excel(excel_save_path, index=False, header=False)
                status_label.config(text="File converted to Excel successfully")
                clear_uploads_folder()
            else:
                status_label.config(text="Excel save cancelled")
        else:
            status_label.config(text="File not found")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")
        status_label.config(text="Failed to convert file")


def upload_file():
    file_path = filedialog.askopenfilename()
    if file_path and os.path.exists(file_path):
        filename = os.path.basename(file_path)
        if allowed_file(filename):
            destination_path = os.path.join(UPLOAD_FOLDER, filename)
            shutil.copy(file_path, destination_path)
            upload_button.pack_forget()
            app.file_path = destination_path
            status_label.config(text="File uploaded successfully")
            continue_button.pack()
            row_count_button.pack(pady=5)
        else:
            status_label.config(text="Invalid file format")
    else:
        status_label.config(text="File not found")

def display_row_count():
    if app.file_path and os.path.exists(app.file_path):
        try:
            file_path = os.path.join(UPLOAD_FOLDER, os.path.basename(app.file_path))  # Get full file path in uploads folder
            total_rows = get_total_rows(file_path)
            if total_rows is not None:
                status_label.config(text=f"File has {total_rows} rows")
            else:
                status_label.config(text="Error: Unable to get row count")
        except Exception as e:
            print(f"An error occurred while getting row count: {e}")
            status_label.config(text="Error: Unable to get row count")
    else:
        status_label.config(text="File path is not valid")




def upload_text_file():
    file_path = filedialog.askopenfilename()
    if file_path and os.path.exists(file_path):
        filename = os.path.basename(file_path)
        if allowed_file(filename):
            destination_path = os.path.join(UPLOAD_FOLDER, filename)
            shutil.copy(file_path, destination_path)
            status_label.config(text="Txt file Uploaded Successfully")
            app.file_path = destination_path
            txt_to_excel_file(destination_path)
        else:
            status_label.config(text="Invalid file format")
    else:
        status_label.config(text="File not found")


def clear_uploads_folder():
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file: {e}")


def align_script(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_csv(file_path, sep='\t', header=None)

        # Replace '2022FRIB' in the DataFrame with None in the appropriate column
        df.loc[df[3] == '2022FRIB', 3] = None

        # Convert the 'Date' column to the desired format
        df[4] = df[4].astype(str).str[:8]

        # Save the modified DataFrame back to the Excel file
        df.to_csv(file_path, sep='\t', index=False, header=False)
        
        return file_path
        
    except Exception as e:
        print(f"An error occurred during alignment: {e}")
        return None


def run_macro(file_path):
    try:        
        align_script(file_path)
        wb = xw.Book(file_path)
        sheet = wb.sheets[0]

        # Clear existing headers and set new headers
        headers = ["Person ID", "PERNR", "Sub Account", "Project", "Date", "hours", "wo", "REGU"]
        sheet.range("1:1").clear_contents()
        for i, header in enumerate(headers, start=1):
            sheet.cells(1, i).value = header

        # Autofit columns
        sheet.autofit()

        # Close the workbook
        wb.save()
        wb.close()

        return file_path

    except Exception as e:
        print(f"An error occurred while modifying the file: {e}")
        return None


def continue_process():
    if app.file_path and os.path.exists(app.file_path):
        modified_file_path = run_macro(app.file_path)
        if modified_file_path:
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Save Modified File As")
            if save_path:
                shutil.copy(modified_file_path, save_path)
                status_label.config(text=f"File modified successfully and saved as {save_path}")
                clear_uploads_folder()
            else:
                status_label.config(text="Save cancelled")
        else:
            status_label.config(text="Failed to modify file")
    else:
        status_label.config(text="File path is not valid")


header_label = tk.Label(app, text="FRIB Payroll Validation", font=("Arial", 16, "bold"))
header_label.pack(pady=10)

upload_frame = tk.Frame(app)
upload_frame.pack(pady=10)

upload_label = tk.Label(upload_frame, text="Upload Excel File")
upload_label.grid(row=0, column=0, padx=10)

upload_button = tk.Button(upload_frame, text="Browse", command=upload_file)
upload_button.grid(row=0, column=1, padx=10)

status_label = tk.Label(app, text="")
status_label.pack(pady=5)

continue_button = tk.Button(app, text="Continue", command=continue_process)
continue_button.pack_forget()

txt_to_excel_button = tk.Button(app, text="Convert Text to Excel", command=upload_text_file)
txt_to_excel_button.pack(pady=5)

progress = ttk.Progressbar(app, orient="horizontal", length=300, mode="determinate")
progress.pack_forget()

percentage_label = tk.Label(app, text="")
percentage_label.pack(pady=5)
percentage_label.pack_forget()

row_count_button = tk.Button(app, text="Display Row Count", command=display_row_count)
row_count_button.pack_forget()


app.file_path = None

if __name__ == "__main__":
    app.mainloop()
