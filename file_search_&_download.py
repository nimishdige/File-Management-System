import tkinter as tk
from tkinter import messagebox, filedialog
import os
import shutil
import pandas as pd

class FileSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Search and Download")
        self.root.geometry("550x400")
        self.root.configure(bg="#f0f0f0")

        # Main frame
        self.main_frame = tk.Frame(root, bg="#f0f0f0")
        self.main_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Year selection dropdown
        self.year_label = tk.Label(self.main_frame, text="Select Year:", bg="#f0f0f0", font=("Helvetica", 12))
        self.year_label.grid(row=0, column=0, pady=(0, 5), sticky="w")

        self.selected_year = tk.StringVar(root)
        self.year_dropdown = tk.OptionMenu(self.main_frame, self.selected_year, "")
        self.year_dropdown.grid(row=0, column=1, pady=(0, 5), padx=(0, 20))

        # Fetch folders from IP folder
        ip_folder_path = 'C:/Users/ITADMIN/Documents/Nimesh Dige/IP'
        if os.path.exists(ip_folder_path):
            self.year_options = [folder for folder in os.listdir(ip_folder_path) if os.path.isdir(os.path.join(ip_folder_path, folder))]
            self.selected_year.set(self.year_options[0] if self.year_options else "")
            self.year_menu = self.year_dropdown['menu']
            self.year_menu.delete(0, 'end')
            for year in self.year_options:
                self.year_menu.add_command(label=year, command=tk._setit(self.selected_year, year))

        # File name entry
        self.file_name_label = tk.Label(self.main_frame, text="Enter File Name:", bg="#f0f0f0", font=("Helvetica", 12))
        self.file_name_label.grid(row=1, column=0, pady=(0, 5), sticky="w")

        self.file_name_entry = tk.Entry(self.main_frame, font=("Helvetica", 12), width=20)
        self.file_name_entry.grid(row=1, column=1, pady=(0, 5), padx=(0, 20))
        self.file_name_entry.bind("<KeyRelease>", self.show_suggestions)

        # Suggestions Listbox with Scrollbar
        self.suggestions_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.suggestions_frame.grid(row=2, column=0, columnspan=5000, pady=(10, 5), padx=(0, 20))

        self.scrollbar = tk.Scrollbar(self.suggestions_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.suggestions_listbox = tk.Listbox(self.suggestions_frame, font=("Helvetica", 10), selectmode=tk.MULTIPLE,
                                              yscrollcommand=self.scrollbar.set, width=70)
        self.suggestions_listbox.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        self.scrollbar.config(command=self.suggestions_listbox.yview)

        # Download Selected Files button
        self.download_button = tk.Button(self.main_frame, text="Download Selected Files", bg="skyblue", fg="white",
                                         font=("Helvetica", 12, "bold"), command=self.download_selected_files)
        self.download_button.grid(row=3, column=0, pady=(20, 10), sticky="w")

        # Download Bulk button
        self.download_bulk_button = tk.Button(self.main_frame, text="Download Bulk", bg="#4CAF50", fg="white",
                                              font=("Helvetica", 12, "bold"), command=self.download_bulk)
        self.download_bulk_button.grid(row=3, column=1, pady=(20, 10), padx=(10, 0), sticky="e")

        # Download file path button with download icon
#         self.download_icon = tk.PhotoImage(file='C:/Users/ITADMIN/Documents/download.png')
        self.download_file_path_button = tk.Button(self.main_frame, text="Records", bg="lightgrey", command=self.download_file_path)
        self.download_file_path_button.grid(row=0, column=2, pady=(0, 5), padx=(10, 0), sticky="e")

        # Default file path
        self.file_path = "C:/Users/ITADMIN/Documents/new_record_file_saved.xlsx"

    def show_suggestions(self, event=None):
        self.suggestions_listbox.delete(0, tk.END)
        partial_file_name = self.file_name_entry.get()

        if not partial_file_name:
            return

        folder_path = os.path.join("C:/Users/ITADMIN/Documents/Nimesh Dige/IP", self.selected_year.get())

        if not os.path.exists(folder_path):
            messagebox.showerror("Error", f"No folder found for the specified year.")
            return

        matching_files = [file for file in os.listdir(folder_path) if partial_file_name.lower() in file.lower()]

        for file in matching_files:
            self.suggestions_listbox.insert(tk.END, file)

    def download_selected_files(self):
        year = self.selected_year.get()

        if not year.isdigit():
            messagebox.showerror("Error", "Please enter a valid year.")
            return

        folder_path = os.path.join("C:/Users/ITADMIN/Documents/Nimesh Dige/IP", year)

        if not os.path.exists(folder_path):
            messagebox.showerror("Error", f"No folder found for the year {year}.")
            return

        selected_indices = self.suggestions_listbox.curselection()

        if not selected_indices:
            messagebox.showerror("Error", "Please select one or more files to download.")
            return

        selected_files = [self.suggestions_listbox.get(index) for index in selected_indices]

        try:
            downloaded_files_count = 0
            destination_directory = filedialog.askdirectory()
            for file_name in selected_files:
                file_path = os.path.join(folder_path, file_name)
                destination_path = os.path.join(destination_directory, file_name)
                shutil.copy(file_path, destination_path)
                downloaded_files_count += 1

            messagebox.showinfo("Download", f"{downloaded_files_count} files have been downloaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download files: {str(e)}")

    def download_bulk(self):
        try:
            downloaded_files_count = 0
            destination_directory = filedialog.askdirectory()
            bulk_file = pd.read_excel(self.file_path, dtype=str)

            main_dir = 'C:/Users/ITADMIN/Documents/Nimesh Dige/IP'

            for i in bulk_file['File_name']:
                year = i.split("_")[-3].split(".")[0]
                destination_path = os.path.join(destination_directory, year)
                os.makedirs(destination_path, exist_ok=True)
                dircopy = os.path.join(main_dir, year, i)
                new_file_path = os.path.join(destination_path, i)
                shutil.copy(dircopy, new_file_path)
                downloaded_files_count += 1

            messagebox.showinfo("Download", f"{downloaded_files_count} files have been downloaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download files: {str(e)}")

    def download_file_path(self):
        destination_directory = filedialog.askdirectory()
        if destination_directory:
            try:
                shutil.copy(self.file_path, destination_directory)
                messagebox.showinfo("Download", "File downloaded successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download file: {str(e)}")
        else:
            messagebox.showinfo("Download", "No destination directory selected.")

root = tk.Tk()
app = FileSearchApp(root)
root.mainloop()
