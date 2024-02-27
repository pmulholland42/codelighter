import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def select_xlsx_file():
    # Define the file types that should be allowed, in this case, .xlsx files
    filetypes = [('Excel files', '*.xlsx')]
    
    # Open the file dialog and store the selected file path
    filepath = filedialog.askopenfilename(title='Open a file', initialdir='/', filetypes=filetypes)
    
    if filepath:
        print(f"Selected file: {filepath}")
        # Perform your operations with the selected file here
        try:
          wb = load_workbook(filename = filepath)
          print(wb.sheetnames)
        except:
            print("Couldn't find . Make sure you run this program in the same folder as the excel file.")
    else:
        print("No file selected.")

# Create the main window
root = tk.Tk()
root.title('Select an .xlsx File')
root.geometry('300x150')  # Width x Height

# Create a button that will open the file dialog when clicked
select_file_btn = tk.Button(root, text='Select .xlsx File', command=select_xlsx_file)
select_file_btn.pack(expand=True)

# Start the Tkinter event loop
root.mainloop()
