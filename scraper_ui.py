import tkinter as tk
from tkinter import filedialog
import os
import sys
sys.path.append(os.path.abspath(r"C:\Users\40gil\Desktop\Helpful\Scraping"))
from LivneScraperPY import main

def browse_file(entry):
    filename = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def browse_folder(entry):
    folder = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder)
def on_submit():
    company_name = company_name_input.get()
    bond_name = bond_name_input.get()
    to_save_path = to_save_path_input.get()
    chrome_driver_path = chrome_driver_path_input.get()

    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, "...התחלתי לעבוד\n")
    output_text.update_idletasks()  # Force GUI update

    try:
        main.start(_company_name=company_name, bond_name=bond_name,
                      to_save_path=to_save_path, chrome_driver=chrome_driver_path)
        output_text.insert(tk.END, "זהו. סיכום החברה נוצר ונשמר במקום הרצוי \n")
    except Exception as e:
        error_message = f"Error: {str(e)}\nPlease check your internet connection and ensure the entered information is correct."
        output_text.insert(tk.END, error_message)
    output_text.update_idletasks()

# Create the main window
root = tk.Tk()
root.title("HelpfulPro Scrapper")

# Text Input 1
company_name_input_text = tk.Label(root, text="שם החברה :")
company_name_input_text.grid(row=0, column=0, padx=10, pady=10)
company_name_input = tk.Entry(root)
company_name_input.grid(row=0, column=1, padx=10, pady=10)

# Text Input 2
bond_name_input_text = tk.Label(root, text=f'שם האגח בביזפורטל: \n "לדוגמה, עבור "פועלים אגח 100" יש להכניס "פועלים אגח ')
bond_name_input_text.grid(row=1, column=0, padx=10, pady=10)
bond_name_input = tk.Entry(root)
bond_name_input.grid(row=1, column=1, padx=10, pady=10)

# Path 1
to_save_path_input_text = tk.Label(root, text="? איפה לשמור את הקובץ :")
to_save_path_input_text.grid(row=2, column=0, padx=10, pady=10)
to_save_path_input = tk.Entry(root)
to_save_path_input.grid(row=2, column=1, padx=10, pady=10)
button_browse_1 = tk.Button(root, text="עיון...", command=lambda: browse_folder(to_save_path_input))
button_browse_1.grid(row=2, column=2, padx=10, pady=10)

# Path 2
chrome_driver_path_input_text = tk.Label(root, text="Crome driver:")
chrome_driver_path_input_text.grid(row=3, column=0, padx=10, pady=10)
chrome_driver_path_input = tk.Entry(root)
chrome_driver_path_input.grid(row=3, column=1, padx=10, pady=10)
button_browse_2 = tk.Button(root, text="עיון...", command=lambda: browse_file(chrome_driver_path_input))
button_browse_2.grid(row=3, column=2, padx=10, pady=10)

output_text_label = tk.Label(root, text="Output:")
output_text_label.grid(row=5, column=0, padx=10, pady=10, sticky="e")
output_text = tk.Text(root, height=5, width=50)
output_text.grid(row=5, column=1, columnspan=2, padx=10, pady=10)

# Submit Button
submit_button = tk.Button(root, text="התחל", command=on_submit)
submit_button.grid(row=4, column=0, columnspan=3, pady=20)

# Run the main loop
root.mainloop()
