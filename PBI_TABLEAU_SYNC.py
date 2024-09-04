import tkinter as tk
from tkinter import filedialog
import customtkinter
import pandas as pd
from pyxlsb import open_workbook
from openpyxl import load_workbook
import os

def selectExcelFile(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def add_network_column_and_filter(pbi_raw_path, tableau_raw_path, pbi_prev_month_path, output_path):
    # Load the PBI raw file
    pbi_df = pd.read_excel(pbi_raw_path)
    print("PBI Raw Columns:", pbi_df.columns.tolist())  # Debugging: print column names
    
    # Load the Tableau raw file
    tableau_df = pd.read_excel(tableau_raw_path)
    print("Tableau Raw Columns:", tableau_df.columns.tolist())  # Debugging: print column names
    
    # Load the PBI previous month file
    pbi_prev_month_df = pd.read_excel(pbi_prev_month_path)
    print("PBI Previous Month Columns:", pbi_prev_month_df.columns.tolist())  # Debugging: print column names
    
    # Filter out rows where the 'Network' column has the value 'AFS Shuttle' in Tableau raw
    tableau_df_filtered = tableau_df[tableau_df['Network'] != 'AFS Shuttle']
    
    # Merge the PBI raw DataFrame with the filtered Tableau raw DataFrame
    # based on matching 'TransportOrderId' in PBI raw with 'Transportorder Id (Transportorder)' in Tableau raw
    merged_df = pd.merge(
        pbi_df,
        tableau_df_filtered[['Transportorder Id (Transportorder)', 'Network']],
        how='left',
        left_on='TransportOrderId',
        right_on='Transportorder Id (Transportorder)'
    )
    
    print("Merged DataFrame Columns:", merged_df.columns.tolist())  # Debugging: print column names after merge
    
    # Drop the unnecessary 'Transportorder Id (Transportorder)' column
    if 'Transportorder Id (Transportorder)' in merged_df.columns:
        merged_df.drop(columns=['Transportorder Id (Transportorder)'], inplace=True)
    
    # Rename 'Network_y' to 'Network' and drop 'Network_x' if they exist
    if 'Network_y' in merged_df.columns:
        merged_df.rename(columns={'Network_y': 'Network'}, inplace=True)
        if 'Network_x' in merged_df.columns:
            merged_df.drop(columns=['Network_x'], inplace=True)
    
    # Filter out rows with 'Network' equal to 'AFS Shuttle' again in the merged DataFrame
    merged_df_filtered = merged_df[merged_df['Network'] != 'AFS Shuttle']

    # Remove duplicate rows based on all columns
    merged_df_filtered.drop_duplicates(inplace=True)
    
    # Save both the filtered Tableau DataFrame, the merged DataFrame, and the PBI previous month into different sheets of the same Excel file
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged_df_filtered.to_excel(writer, index=False, sheet_name='PBI_Raw_Updated')
            tableau_df_filtered.to_excel(writer, index=False, sheet_name='Tableau_Raw_Filtered')
            pbi_prev_month_df.to_excel(writer, index=False, sheet_name='PBI_Previous_Month')
        print(f"Updated file saved as {output_path}")
    except PermissionError:
        print(f"Permission denied: Unable to save the file. Please close '{output_path}' if it is open in another program or choose a different output path.")

def upload_files(pathEntry1, pathEntry2, pathEntry3):
    # Get the file paths from the entries
    pbi_raw_path = pathEntry1.get()
    tableau_raw_path = pathEntry2.get()
    pbi_prev_month_path = pathEntry3.get()
    
    if pbi_raw_path and tableau_raw_path and pbi_prev_month_path:  # Check if all file paths are not empty
        # Set the output path for the new Excel file
        output_path = os.path.splitext(pbi_raw_path)[0] + "_updated.xlsx"
        add_network_column_and_filter(pbi_raw_path, tableau_raw_path, pbi_prev_month_path, output_path)
    else:
        print("Please select the PBI raw, Tableau raw, and PBI previous month files.")

def app():
    customtkinter.set_appearance_mode("System")
    
    app = customtkinter.CTk()
    app.title("Tariff extractor")

    frame = customtkinter.CTkFrame(app)
    frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)

    title1 = customtkinter.CTkLabel(frame, text="PBI raw")
    title1.grid(row=0, column=0, pady=10, padx=5)

    pathEntry1 = customtkinter.CTkEntry(frame)
    pathEntry1.grid(row=0, column=1, pady=10, padx=5)

    browseButton1 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        command=lambda: selectExcelFile(pathEntry1)
    )
    browseButton1.grid(row=0, column=2, pady=10, padx=5)

    title2 = customtkinter.CTkLabel(frame, text="Tableau raw")
    title2.grid(row=1, column=0, pady=10, padx=5)

    pathEntry2 = customtkinter.CTkEntry(frame)
    pathEntry2.grid(row=1, column=1, pady=10, padx=5)

    browseButton2 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        command=lambda: selectExcelFile(pathEntry2)
    )
    browseButton2.grid(row=1, column=2, pady=10, padx=5)

    title3 = customtkinter.CTkLabel(frame, text="PBI previous month")
    title3.grid(row=2, column=0, pady=10, padx=5)

    pathEntry3 = customtkinter.CTkEntry(frame)
    pathEntry3.grid(row=2, column=1, pady=10, padx=5)

    browseButton3 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        command=lambda: selectExcelFile(pathEntry3)
    )
    browseButton3.grid(row=2, column=2, pady=10, padx=5)

    uploadButton = customtkinter.CTkButton(
        frame,
        text="Upload",
        command=lambda: upload_files(pathEntry1, pathEntry2, pathEntry3)
    )
    uploadButton.grid(row=3, column=1, pady=10, padx=5)

    app.mainloop()

app()
