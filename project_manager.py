import pandas as pd
from tkinter import *
from tkinter import ttk
import os
from datetime import datetime

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define the file path relative to the script's directory
file_path = os.path.join(script_dir, "project_db.xlsx")

def excel_init(file_path):
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        # Drop any unnamed columns that may have been added accidentally
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        return df
    return pd.DataFrame(columns=["Project Name", "Date", "User"])

# Saves the data in order in the excel and changes the result label to "Data Saved"
def save_data():
    df = excel_init(file_path)
    
    # Get the selected date as a string
    date_str = f"{year_combobox.get()}-{month_combobox.get()}-{day_combobox.get()}"
    
    new_data = {
        "Project Name": project_name_entry.get(),
        "Date": date_str,
        "User": name_entry.get()
    }
    
    # Ensure that the new entry is not empty or a duplicate
    if new_data["Project Name"] and new_data["User"] and new_data["Date"]:
        # Check if this entry already exists in the DataFrame
        if not ((df["Project Name"] == new_data["Project Name"]) & 
                (df["Date"] == new_data["Date"]) & 
                (df["User"] == new_data["User"])).any():
            
            # Add new data to the DataFrame
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            
            # Save the DataFrame to Excel, excluding NaN values
            df.dropna(how='all', inplace=True)
            df.to_excel(file_path, index=False)
            result_label.config(text="Data Saved")
        else:
            result_label.config(text="Duplicate entry, not saved")
    else:
        result_label.config(text="Incomplete data, not saved")
    
    update_table()

# Updates the Treeview with the latest data from the Excel file
def update_table():
    # Clear the Treeview
    for row in tree.get_children():
        tree.delete(row)
    
    # Read data from Excel
    df = excel_init(file_path)
    print("Data being loaded into the table:")  # Debug statement
    print(df)  # Debug statement
    
    # Insert data into Treeview
    for index, row in df.iterrows():
        tree.insert("", "end", values=(row["Project Name"], row["Date"], row["User"]))

root = Tk()
root.geometry("500x500")

# Get today's date and 10 years in the future
today = datetime.today()
current_year = today.year
current_month = today.month
current_day = today.day
future_year = current_year + 10

Label(root, text="Name:").pack(pady=5)
name_entry = Entry(root)
name_entry.pack(pady=5)

Label(root, text="Project name:").pack(pady=5)
project_name_entry = Entry(root)
project_name_entry.pack(pady=5)

# Dropdowns for selecting month, day, and year
Label(root, text="Select Date:").pack(pady=5)

# Month dropdown
month_combobox = ttk.Combobox(root, values=[f"{i:02d}" for i in range(1, 13)], state="readonly")
month_combobox.set(f"{current_month:02d}")
month_combobox.pack(pady=5)

# Day dropdown
day_combobox = ttk.Combobox(root, values=[f"{i:02d}" for i in range(1, 32)], state="readonly")
day_combobox.set(f"{current_day:02d}")
day_combobox.pack(pady=5)

# Year dropdown
year_combobox = ttk.Combobox(root, values=[str(i) for i in range(current_year, future_year + 1)], state="readonly")
year_combobox.set(str(current_year))
year_combobox.pack(pady=5)

# After data is saved and excel is changed, it will show prompt, "Data Saved"
Button(root, text="Save Data", command=save_data).pack(pady=10)
result_label = Label(root, text="")
result_label.pack(pady=5)

# Treeview to display the data
tree = ttk.Treeview(root, columns=("Project Name", "Date", "User"), show='headings')
tree.heading("Project Name", text="Project Name")
tree.heading("Date", text="Date")
tree.heading("User", text="User")
tree.pack(pady=20, fill="x")

# Initialize table with current data
update_table()
root.mainloop()
