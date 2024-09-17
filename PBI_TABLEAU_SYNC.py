import tkinter as tk
import datetime
from tkinter import filedialog
import customtkinter
import pandas as pd
from openpyxl import load_workbook
import os
import tableauserverclient as TSC
from io import StringIO

class TableauWorkbookDownloader:
    def __init__(self, server_url, personal_access_token_name, personal_access_token_secret):
        self.server_url = server_url
        self.personal_access_token_name = personal_access_token_name
        self.personal_access_token_secret = personal_access_token_secret
        self.server = None

    def connect_to_server(self):
        if not self.server:  # Connect to the server only if not already connected
            print("Connecting to Tableau Server...")
            self.server = TSC.Server(self.server_url, use_server_version=True)
            tableau_auth = TSC.PersonalAccessTokenAuth(self.personal_access_token_name, self.personal_access_token_secret)
            self.server.auth.sign_in(tableau_auth)
            print("Connected to Tableau Server.")

    def find_workbook_by_name_and_id(self, workbook_name, workbook_id):
        print(f"Searching for workbook '{workbook_name}' with ID '{workbook_id}'...")
        req_option = TSC.RequestOptions()
        req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name, TSC.RequestOptions.Operator.Equals, workbook_name))

        all_workbooks = []
        req_option.page_size = 100
        req_option.page_number = 1

        while True:
            workbooks, pagination_item = self.server.workbooks.get(req_option)
            all_workbooks.extend(workbooks)

            if not pagination_item or pagination_item.page_number * req_option.page_size >= pagination_item.total_available:
                break

            req_option.page_number += 1

        matching_workbooks = [wb for wb in all_workbooks if wb.id == workbook_id]
        print(f"Found {len(matching_workbooks)} matching workbooks.")

        return matching_workbooks

    def download_view_as_dataframe(self, workbook_name, workbook_id, view_name, filters=None, batch_size=50):
        print(f"Starting download for view '{view_name}' in workbook '{workbook_name}'...")
        self.connect_to_server()

        matching_workbooks = self.find_workbook_by_name_and_id(workbook_name, workbook_id)
        if not matching_workbooks:
            print(f"Workbook '{workbook_name}' with ID '{workbook_id}' not found.")
            return None

        workbook = matching_workbooks[0]

        print(f"Populating views for workbook '{workbook_name}'...")
        self.server.workbooks.populate_views(workbook)
        view = next((v for v in workbook.views if v.name == view_name), None)
        if not view:
            print(f"View '{view_name}' not found in workbook '{workbook_name}'.")
            return None

        all_batches = []  # List to store all data frames
        tableau_columns_order = None  # Placeholder for the correct column order

        # Ensure unique filter values and split into batches if necessary
        if filters and isinstance(filters, dict):
            print(f"Applying filters to view '{view_name}'...")
            for key, values in filters.items():
                unique_values = list(set(values))
                print(f"Found {len(unique_values)} unique filter values for key '{key}'.")

                for i in range(0, len(unique_values), batch_size):
                    batch_values = unique_values[i:i + batch_size]
                    csv_req_options = TSC.CSVRequestOptions()
                    csv_req_options.vf(key, ",".join(map(str, batch_values)))
                    self.server.views.populate_csv(view, csv_req_options)

                    try:
                        raw_csv_data = b''.join(view.csv).decode('utf-8')
                        batch_data = pd.read_csv(StringIO(raw_csv_data), sep=',', low_memory=False)

                        if tableau_columns_order is None:
                            tableau_columns_order = list(batch_data.columns)

                        batch_data = batch_data[tableau_columns_order]
                        all_batches.append(batch_data)
                    except Exception as e:
                        print(f"Error reading CSV data: {e}")
                        return None
        else:
            print(f"No filters applied. Downloading view '{view_name}' without filters...")
            csv_req_options = TSC.CSVRequestOptions()
            self.server.views.populate_csv(view, csv_req_options)

            try:
                raw_csv_data = b''.join(view.csv).decode('utf-8')
                batch_data = pd.read_csv(StringIO(raw_csv_data), sep=',', low_memory=False)
                tableau_columns_order = list(batch_data.columns)
                batch_data = batch_data[tableau_columns_order]
                all_batches.append(batch_data)
            except Exception as e:
                print(f"Error reading CSV data: {e}")
                return None

        print(f"Combining all downloaded batches into a single DataFrame...")
        combined_data = pd.concat(all_batches, ignore_index=True)
        return combined_data

def selectExcelFile(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def add_network_column_and_filter(pbi_raw_path, pbi_prev_month_path, output_path):
    server_url = "https://tableau.analytics.4flow-vista.com"
    personal_access_token_name = "VS code 2"
    personal_access_token_secret = "/9xmYHc7SLGm4stLCDfzRQ==:iplVfimCZn4t8hWczjxCB74eADOmNuHw"
    
    workbook_name_tableauraw = "EMS Invoicing"
    workbook_id_tableauraw = "ee931da9-14fc-4ac3-ad02-3aec44f17f54"
    view_name_tableauraw = "TO Count"
    
    workbook_name_ex = "AFS EMS Raw Data & Overview"
    workbook_id_ex = "c701e7de-f48d-41ae-9a3e-aff0afd2bcbf"
    view_name_ex = "Raw Last Month"
    
    tableau_downloader = TableauWorkbookDownloader(server_url, personal_access_token_name, personal_access_token_secret)
    tableau_downloader.connect_to_server()
    
    # Load the PBI raw file
    pbi_df = pd.read_excel(pbi_raw_path)
    print("PBI Raw Columns:", pbi_df.columns.tolist())
    
    # Load the Tableau raw file
    tableau_df = tableau_downloader.download_view_as_dataframe(workbook_name_tableauraw, workbook_id_tableauraw, view_name_tableauraw)
    if tableau_df is None:
        print(f"Error: Workbook '{workbook_name_tableauraw}' or view '{view_name_tableauraw}' not found.")
        return
    
    print("Tableau Raw Columns:", tableau_df.columns.tolist())
    
    # Load the PBI previous month file
    pbi_prev_month_df = pd.read_excel(pbi_prev_month_path)
    print("PBI Previous Month Columns:", pbi_prev_month_df.columns.tolist())
    
    # Load the Tableau exception workbook data
    tableau_ex_df = tableau_downloader.download_view_as_dataframe(workbook_name_ex, workbook_id_ex, view_name_ex)
    if tableau_ex_df is None:
        print(f"Error: Exception workbook '{workbook_name_ex}' or view '{view_name_ex}' not found.")
        return
    
    print("Tableau Exception Columns:", tableau_ex_df.columns.tolist())

    # Remove duplicates from the exception data based on 'Issue ID w/o duplicates'
    tableau_ex_filtered_df = tableau_ex_df.drop_duplicates(subset=['Issue ID w/o duplicates'], keep='first')

    # Remove the last row from both unfiltered and filtered exception DataFrames
    tableau_ex_df_no_last = tableau_ex_df.iloc[:-1]
    tableau_ex_filtered_df_no_last = tableau_ex_filtered_df.iloc[:-1]
    
    print("Filtered Tableau Exception Data Columns:", tableau_ex_filtered_df_no_last.columns.tolist())
    
    # Merge the PBI raw DataFrame with the Tableau raw DataFrame
    merged_df = pd.merge(
        pbi_df,
        tableau_df[['Transportorder Id (Transportorder)', 'Network']],
        how='left',
        left_on='TransportOrderId',
        right_on='Transportorder Id (Transportorder)'
    )
    
    print("Merged DataFrame Columns:", merged_df.columns.tolist())
    
    # Drop the unnecessary 'Transportorder Id (Transportorder)' column
    if 'Transportorder Id (Transportorder)' in merged_df.columns:
        merged_df.drop(columns=['Transportorder Id (Transportorder)'], inplace=True)
    
    # Rename 'Network_y' to 'Network' and drop 'Network_x' if they exist
    if 'Network_y' in merged_df.columns:
        merged_df.rename(columns={'Network_y': 'Network'}, inplace=True)
        if 'Network_x' in merged_df.columns:
            merged_df.drop(columns=['Network_x'], inplace=True)
    
    # Filter out 'AFS Shuttle' for further analysis
    tableau_df_filtered = tableau_df[tableau_df['Network'] != 'AFS Shuttle']
    pbi_df_filtered = merged_df[merged_df['Network'] != 'AFS Shuttle']
    
    # Identify missing transport orders in Tableau but present in PBI
    missing_in_tableau = pbi_df_filtered[~pbi_df_filtered['TransportOrderId'].isin(tableau_df_filtered['Transportorder Id (Transportorder)'])]
    
    # Identify missing transport orders in PBI but present in Tableau
    missing_in_pbi = tableau_df_filtered[~tableau_df_filtered['Transportorder Id (Transportorder)'].isin(pbi_df_filtered['TransportOrderId'])]
    
    # Identify duplicate transport orders in PBI raw and PBI previous month
    duplicates_pbi = pd.concat([pbi_df, pbi_prev_month_df]).duplicated(subset=['TransportOrderId'], keep=False)
    duplicates_pbi_df = pd.concat([pbi_df, pbi_prev_month_df])[duplicates_pbi]
    # Now calculate the counts after merging, as Network is added after merge
    pbi_to_count = len(merged_df)  # Total rows in merged PBI data
    shuttles_count = len(merged_df[merged_df['Network'] == 'AFS Shuttle'])  # Rows with 'AFS Shuttle'
    new_to_count = pbi_to_count - shuttles_count  # New TO Count
    duplicates_count = len(duplicates_pbi_df)  # Duplicates count
    new_new_to_count = new_to_count - duplicates_count  # New New TO Count
    exceptions_count = len(tableau_ex_filtered_df_no_last)  # Exceptions count
    exc_rate = exceptions_count / new_new_to_count if new_new_to_count > 0 else 0  # Exception rate
    formatted_rate = "{:,.2f}%".format(exc_rate * 100)   # Prepare summary data
    summary_data = {
        'Metric': [
            'PBI TO Count',
            'Shuttles',
            'New TO Count',
            'Duplicates',
            'New New TO Count',
            'Exceptions Count',
            'Exc Rate'
        ],
        'Value': [
            pbi_to_count,
            shuttles_count,
            new_to_count,
            duplicates_count,
            new_new_to_count,
            exceptions_count,
            formatted_rate
        ]
    }

    summary_df = pd.DataFrame(summary_data)
    # Save all DataFrames into different sheets of the same Excel file
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
            
            # Save the PBI raw updated data
            merged_df.to_excel(writer, index=False, sheet_name='PBI_Raw_Updated')
            
            # Save the unfiltered Tableau raw data
            tableau_df.to_excel(writer, index=False, sheet_name='Tableau_Raw_Unfiltered')
            
           
            
            # Save the PBI previous month data
            pbi_prev_month_df.to_excel(writer, index=False, sheet_name='PBI_Previous_Month')
            
            # Save the missing transport orders in Tableau
            missing_in_tableau.to_excel(writer, index=False, sheet_name='Missing_In_Tableau')
            
            # Save the missing transport orders in PBI
            missing_in_pbi.to_excel(writer, index=False, sheet_name='Missing_In_PBI')
            
            # Save the duplicate transport orders in PBI
            duplicates_pbi_df.to_excel(writer, index=False, sheet_name='Duplicates_In_PBI')
            
            # Save the Tableau exception data (with last row removed)

            # Save the filtered exception data with removed duplicates and without the last row
            tableau_ex_filtered_df_no_last.to_excel(writer, index=False, sheet_name='Filtered_Exception_Raw')
            
            

        print(f"Updated file saved as {output_path}")
    except PermissionError:
        print(f"Permission denied: Unable to save the file. Please close '{output_path}' if it is open in another program or choose a different output path.")
def get_prefilled_save_path():
    # Get the current date to format the file name
    current_date = datetime.datetime.now()
    month_year = current_date.strftime("%m/%Y")
    
    # Set the default file name with the current month and year
    default_file_name = f"Monthly Invoicing TO/Exceptions {month_year}"
    
    # Replace any slashes in the filename with dashes or other valid characters
    default_file_name = default_file_name.replace("/", "-")
    
    # Open a "Save As" dialog with the default name
    file_path = filedialog.asksaveasfilename(
        initialdir=os.getcwd(),
        title="Save file",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
        initialfile=default_file_name
    )
    
    return file_path

def upload_files(pathEntry1, pathEntry3):
    pbi_raw_path = pathEntry1.get()
    pbi_prev_month_path = pathEntry3.get()
    
    if pbi_raw_path and pbi_prev_month_path:
        # Prompt the user to choose a location to save the file with the prefilled name
        output_path = get_prefilled_save_path()
        
        if output_path:
            add_network_column_and_filter(pbi_raw_path, pbi_prev_month_path, output_path)
        else:
            print("No file path selected. Operation canceled.")
    else:
        print("Please select the PBI raw and PBI previous month files.")
def app():
    customtkinter.set_appearance_mode("System")
    
    app = customtkinter.CTk()
    app.title("Tableau_PBI_Synchronizer")

    frame = customtkinter.CTkFrame(app)
    frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)

    title1 = customtkinter.CTkLabel(frame, text="PBI raw")
    title1.grid(row=0, column=0, pady=10, padx=5)

    pathEntry1 = customtkinter.CTkEntry(frame)
    pathEntry1.grid(row=0, column=1, pady=10, padx=5)

    browseButton1 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        command=lambda: selectExcelFile(pathEntry1),
        fg_color="#EE7203"
    )
    browseButton1.grid(row=0, column=2, pady=10, padx=5)

    title3 = customtkinter.CTkLabel(frame, text="PBI previous month")
    title3.grid(row=2, column=0, pady=10, padx=5)

    pathEntry3 = customtkinter.CTkEntry(frame)
    pathEntry3.grid(row=2, column=1, pady=10, padx=5)

    browseButton3 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        command=lambda: selectExcelFile(pathEntry3),
        fg_color="#EE7203"
    )
    browseButton3.grid(row=2, column=2, pady=10, padx=5)

    uploadButton = customtkinter.CTkButton(
        frame,
        text="Upload",
        command=lambda: upload_files(pathEntry1, pathEntry3),
        fg_color="#EE7203"
    )
    uploadButton.grid(row=3, column=1, pady=10, padx=5)

    app.mainloop()

app()
