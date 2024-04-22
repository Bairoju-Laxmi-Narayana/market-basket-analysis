import tkinter as tk
from tkinter import messagebox
import pandas as pd

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"+{x}+{y}")

def on_configure(event):
    if event.widget == root:  # Check if event is on the root window
        center_window(root)

def search_support():
    input_1 = antecedent_entry.get()
    input_2 = consequent_entry.get()

    try:
        df1 = pd.read_excel('C:\\Users\\91934\\Desktop\\confidence.xlsx')
        df2 = pd.read_excel('C:\\Users\\91934\\Desktop\\rules.xlsx')
    except FileNotFoundError:
        messagebox.showerror("Error", "Excel files not found.")
        return

    support1 = search_support_in_sheet(df1, input_1, input_2)
    support2 = search_support_in_sheet(df2, input_1, input_2)

    result_text = f"Excel Sheet 1 Support: {support1}\nExcel Sheet 2 Support: {support2}"
    result_label.config(text=result_text)

def search_support_in_sheet(df, input_1, input_2):
    match1 = df[(df['antecedents'].str.contains(input_1)) & (df['consequents'].str.contains(input_2))]
    match2 = df[(df['antecedents'].str.contains(input_2)) & (df['consequents'].str.contains(input_1))]

    if not match1.empty:
        return match1.iloc[0]['support']
    elif not match2.empty:
        return match2.iloc[0]['support']
    else:
        return "Not found"

root = tk.Tk()
root.title("Support Value Search")

# Antecedent input
antecedent_label = tk.Label(root, text="Enter Antecedent:")
antecedent_label.grid(row=0, column=0, padx=10, pady=10)
antecedent_entry = tk.Entry(root)
antecedent_entry.grid(row=0, column=1, padx=10, pady=10)

# Consequent input
consequent_label = tk.Label(root, text="Enter Consequent:")
consequent_label.grid(row=1, column=0, padx=10, pady=10)
consequent_entry = tk.Entry(root)
consequent_entry.grid(row=1, column=1, padx=10, pady=10)

# Search button
search_button = tk.Button(root, text="Search", command=search_support)
search_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Result label
result_label = tk.Label(root, text="")
result_label.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Bind configure event to root window
root.bind('<Configure>', on_configure)

# Center the window
center_window(root)

root.mainloop()
