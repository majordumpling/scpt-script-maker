#Imports
import tkinter as tk

from tkinter import filedialog, messagebox
from pathlib import Path

def get_valid_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    while True:
        path = filedialog.askopenfilename(title="Select a file")

        if not path:
            # User cancelled the dialog
            confirm = messagebox.askyesno("Cancel", "No file selected. Do you want to cancel?")
            if confirm:
                return None
            else:
                continue

        file = Path(path)
        if file.exists() and file.is_file():
            return path
        else:
            messagebox.showerror("Invalid File", "The selected file does not exist. Please try again.")

def save_file(script):
    # Hide the root window
    root = tk.Tk()
    root.withdraw()

    # Ask user where to save the file
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        title="Save script as"
    )

    if file_path:  # Only proceed if a file was selected
        with open(file_path, 'w', encoding="utf-8") as file:
            file.write(script)
        print(f"Script saved to: {file_path}")
    else:
        print("Save cancelled.")