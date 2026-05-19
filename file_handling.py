from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import xml.etree.ElementTree as ET
import os
from datetime import datetime

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")

home_directory = os.path.expanduser("~")
desktop_path = os.path.join(home_directory, "Desktop")

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}


def load_file(root):
    file_path = None
    while file_path is None:
        file_path = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")],
        )
        if not file_path:
            retry_dialog_result = messagebox.askretrycancel(
                "No file selected", "Please select a file to proceed."
            )
            if retry_dialog_result:
                continue
            else:
                return None

        try:
            if file_path.endswith(".xlsx"):
                df = pd.read_excel(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded Excel file.")
                return df
            elif file_path.endswith(".xls"):
                try:
                    df = pd.read_excel(file_path, header=0, engine="xlrd")
                    messagebox.showinfo("Success", "Successfully loaded Excel file.")
                    return df
                except Exception as e:
                    conversion_attempt_box = messagebox.askquestion(
                        "Error", f"Failed to load. Attempt converting to .xlsx?"
                    )
                    if conversion_attempt_box == "yes":
                        try:
                            df = parse_xml_data_to_df(file_path)
                            messagebox.showinfo("Success", "Successfully loaded file.")
                            return df
                        except Exception as e:
                            retry_dialog_result = messagebox.askretrycancel(
                                "Error",
                                f"An error occurred while reading the file: {e}. Try again?",
                            )
                            if not retry_dialog_result:
                                break
                            file_path = None
                    else:
                        break
            elif file_path.endswith(".csv"):
                df = pd.read_csv(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded CSV file.")
                return df
            else:
                retry_dialog_result = messagebox.askretrycancel(
                    "Error", f"File not found: {file_path}. Select a different file?"
                )
                if not retry_dialog_result:
                    break
                file_path = None
        except Exception as e:
            retry_dialog_result = messagebox.askretrycancel(
                "Error", f"An error occurred while reading the file: {e}. Try again?"
            )
            if not retry_dialog_result:
                break
            file_path = None
    return None


def parse_xml_data_to_df(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    data = []
    columns = []

    table = root.find(".//ss:Table", NS)

    if table is None:
        raise ValueError("Could not find 'Table' element in sheet.")

    header_row = table.find("ss:Row", NS)
    if header_row is not None:
        for cell in header_row.findall("ss:Cell", NS):
            data_element = cell.find("ss:Data", NS)
            if data_element is not None:
                columns.append(data_element.text)

    for row in table.findall("ss:Row", NS)[1:]:
        row_data = []
        for cell in row.findall("ss:Cell", NS):
            cell_data = cell.find("ss:Data", NS)
            if cell_data is not None:
                row_data.append(cell_data.text)
            else:
                row_data.append(None)
        data.append(row_data)

    df = pd.DataFrame(data, columns=columns, index=None)
    excel_file_name = Path(file_path).stem + ".xlsx"
    excel_save_location = os.path.join(desktop_path, excel_file_name)
    df.to_excel(excel_save_location, index=False)
    print(df)
    return df


def save_and_export(formatted_file, save_path, po_number):
    formatted_file = formatted_file.drop(columns=["needs_manual_category"])

    formatted_file["PO_Num"] = po_number

    excel_name = f"{save_path}.xlsx"
    csv_name = f"{save_path}.csv"

    try:
        with pd.ExcelWriter(excel_name, engine="openpyxl") as writer:
            formatted_file.to_excel(writer, sheet_name="PO", index=False)
            format_cols = [
                "External_ID",
                "ISBN",
                "UPC_Code",
                "Webstore_Image_Name",
            ]
            worksheet = writer.sheets["PO"]
            for col_name in format_cols:
                col_idx = formatted_file.columns.get_loc(col_name) + 1
                for row in worksheet.iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = "0"
        formatted_file.to_csv(csv_name, index=False)
        message = (
            f"Processed Excel file:\n{excel_name}.\n\n"
            f"Processed CSV file:\n{csv_name}."
        )
        return True, message
    except Exception as e:
        return False, None, None, str(e)
