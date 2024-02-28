import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import os
import datetime
import shutil
import pandas as pd

class ExpenseTrackerApp:
    def __init__(self, root):
        
        sample_names = pd.read_excel(r"C:\Users\ITADMIN\Documents\test_sample.xlsx", dtype=str)

        self.root = root
        self.root.title("Expense Tracker")
        
        # Dropdown options
        self.company_names = sample_names['Company_names'].tolist()
        self.expense_heads = sample_names['Expense_heads'].tolist()
        self.vendor_names = sample_names['Vendor_Name'].tolist()
        self.employee_codes = sample_names['Emp_Code'].tolist()
        
        # UI elements
        ttk.Label(root, text="Company Name:").grid(row=1, column=1, padx=10, pady=5, sticky='e')
        self.company_name_var = tk.StringVar()
        self.company_name_dropdown = ttk.Combobox(root, state="readonly", textvariable=self.company_name_var, values=self.company_names)
        self.company_name_dropdown.grid(row=1, column=2, padx=10, pady=5, sticky='w')
        
        ttk.Label(root, text="Expense Head:").grid(row=2, column=1, padx=10, pady=5, sticky='e')
        self.expense_head_var = tk.StringVar()
        self.expense_head_dropdown = ttk.Combobox(root, state="readonly", textvariable=self.expense_head_var, values=self.expense_heads)
        self.expense_head_dropdown.grid(row=2, column=2, padx=10, pady=5, sticky='w')
        
        ttk.Label(root, text="Invoice Date:").grid(row=3, column=1, padx=10, pady=5, sticky='e')
        self.invoice_date_var = tk.StringVar()
        self.invoice_date_picker = DateEntry(root, state="readonly", width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.invoice_date_picker.grid(row=3, column=2, padx=10, pady=5, sticky='w')
        
        ttk.Label(root, text="Vendor Name:").grid(row=4, column=1, padx=10, pady=5, sticky='e')
        self.vendor_name_var = tk.StringVar()
        self.vendor_name_dropdown = ttk.Combobox(root, state="readonly", textvariable=self.vendor_name_var, values=self.vendor_names)
        self.vendor_name_dropdown.grid(row=4, column=2, padx=10, pady=5, sticky='w')
        
        ttk.Label(root, text="Employee Code:").grid(row=5, column=1, padx=10, pady=5, sticky='e')
        self.employee_code_var = tk.StringVar()
        self.employee_code_dropdown = ttk.Combobox(root, state="readonly", textvariable=self.employee_code_var, values=self.employee_codes)
        self.employee_code_dropdown.grid(row=5, column=2, padx=10, pady=5, sticky='w')
        
        ttk.Button(root, text="Upload File", command=self.upload_file).grid(row=6, column=1, columnspan=2, padx=10, pady=10, sticky='we')
   
        # Error label
        self.error_label = ttk.Label(root, text="", foreground="red")
        self.error_label.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        
    def upload_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            try:
                invoice_date = self.invoice_date_picker.get_date()
                current_date = datetime.date.today()
                
                selected_date_str = invoice_date.strftime("%Y_%m_%d")
                current_date_str = current_date.strftime("%Y_%m_%d")
                
                # Create directory for current year if not exists
                current_year = current_date.year
            
                root_path = "C:/Users/ITADMIN/Documents/Nimesh Dige/IP"
                
                # Create directory based on selected options
                directory = os.path.join(root_path, str(current_year))
                os.makedirs(directory, exist_ok=True)
                
                # Extract file name and extension
                ofile_name, file_ext = os.path.splitext(os.path.basename(file_path))
                
                # Generate new file name
                file_name = f"{self.company_name_var.get()}_{self.expense_head_var.get()}_{self.vendor_name_var.get()}_{self.employee_code_var.get()}_{selected_date_str}_{current_date_str}"
                
                # Rename the file based on the concatenated string
                new_file_path = os.path.join(directory, f'{file_name}{file_ext}')
                
                rec_file = pd.read_csv("record_file.csv", dtype=str)
                
                # Create a DataFrame for the new data
                new_data = {
                    "Company_names": self.company_name_var.get(),
                    "Expense_heads": self.expense_head_var.get(),
                    "Vendor_Name": self.vendor_name_var.get(),
                    "Emp_Code": self.employee_code_var.get(),
                    "Voucher_date": selected_date_str,
                    "Current_date": current_date_str,
                    "File_name": f"{file_name}{file_ext}"
                }
                
                # Append the new data to the DataFrame
                rec_file = rec_file.append(new_data, ignore_index=True)
                
                # Save the combined data back to the CSV file
                rec_file.to_csv("record_file.csv", index=False)
                
                # Move the file to the new location with the new name
                shutil.copy(file_path, new_file_path)
                
                messagebox.showinfo("Saved",f"File saved as {new_file_path}")

                self.company_name_var.set("")  # Set to blank
                self.expense_head_var.set("")  # Set to blank
                self.vendor_name_var.set("")   # Set to blank
                self.employee_code_var.set("") # Set to blank
                self.invoice_date_picker.set_date(datetime.date.today())
                
            except Exception as e:
                # Inform user about the error
                messagebox.showinfo("Error",f"An error occurred: {e}")

def main():
    root = tk.Tk()
    app = ExpenseTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()