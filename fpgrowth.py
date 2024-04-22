# 
import tkinter as tk
from tkinter import messagebox
import pandas as pd

# Function to search for support values
def search_support():
    input_1 = antecedent_entry.get()
    input_2 = consequent_entry.get()

    # Load Excel sheets
    try:
        df1 = pd.read_excel('C:\\Users\\91934\\Desktop\\confidence.xlsx')
        df2 = pd.read_excel('C:\\Users\\91934\\Desktop\\rules.xlsx')
        df3 = pd.read_excel('C:\\Users\\91934\\Desktop\\eclat.xlsx')
    except FileNotFoundError:
        messagebox.showerror("Error", "Excel files not found.")
        return

    # Search for inputs in all three sheetse
    support1 = search_support_in_sheet(df1, input_1, input_2)
    support2 = search_support_in_sheet(df2, input_1, input_2)
    support3 = search_support_in_sheet(df3, input_1, input_2, antecedent_col="Product 1", consequent_col="Product 2", support_col="Support")

    # Display results
    result_text = f"APRIORI Support: {support1}\nFP GROWTH Support: {support2}\nECLAT Support: {support3}"
    result_label.config(text=result_text)


def search_support_in_sheet(df, input_1, input_2, antecedent_col="antecedents", consequent_col="consequents", support_col="support"):
    # Search for input 1 in antecedents and input 2 in consequents
    match1 = df[(df[antecedent_col].str.contains(input_1)) & (df[consequent_col].str.contains(input_2))]
    match2 = df[(df[antecedent_col].str.contains(input_2)) & (df[consequent_col].str.contains(input_1))]

    if not match1.empty:
        return match1.iloc[0][support_col]
    elif not match2.empty:
        return match2.iloc[0][support_col]
    else:
        return "Not found"


# Create GUI
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

root.mainloop()
