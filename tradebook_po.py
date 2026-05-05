import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import threading
from dotenv import load_dotenv
import requests as req
from requests.exceptions import HTTPError
from PIL import Image
import shutil
import time
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

# def resource_path(relative_path):
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)

# dotenv_path = resource_path("ISBNdb_API_Key.env")

load_dotenv(dotenv_path="./ISBNdb_API_Key.env")
# load_dotenv(dotenv_path=dotenv_path)
api_key = os.getenv("API_KEY")
print(api_key)

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")

home_directory = os.path.expanduser("~")
desktop_path = os.path.join(home_directory, "Desktop")
dl_folder_name = f"book_covers-{date_string}"
dl_folder_home = os.path.join(desktop_path, dl_folder_name)
resized_images = f"{date_string}-resized_images"

cols_to_delete = [
    "Purchase Order",
    "Item #",
    "Author 2",
    "Units",
    "Discount",
    "Discount Code",
    "Format",
    "Secondary Store Category",
    "BISAC Category Code",
    "Publisher Code",
    "SAN",
    "Markup Tags",
    "UPC",
    "Number of Pages",
    "Dimensions",
    "Links",
]

cols_dtypes = {
    "ISBN 10": str,
    "EAN": str,
    "Title": str,
    "Subtitle": str,
    "List Price": float,
    "Cost": float,
    "Edition": str,
    "Print Run": str,
}

vendor_dict = {
    "Penguin Random House": "Penguin Random House LLC",
    "Simon & Schuster": "Simon & Schuster",
    "MPS": "MPS",
    "HC": "Harper Collins Publishers",
    "IN": "Ingram Publisher Services",
    "TW": "Hachette Book Group USA",
    "NO": "W W Norton & Company Inc",
    "394": "Penguin Random House LLC",
    "WI": "John Wiley & Songs Inc",
    "JHU": "Indiana University - Slavica Publishers",
    "ABRA": "Hachette Book Group USA",
    "CM": "Ingram Publisher Services",
    "LCQ": "Harper Collins Publishers",
    "HARB": "W W Norton & Company Inc",
    "DMMQ": "Longleaf Services, Inc",
    "CUNW": "Ingram Publisher Services",
    "BKVR": "Taylor & Francis Group, LLC",
    "REDB": "Red Wheel Weiser, LLC",
    "UOCP": "Ingram Publisher Services",
    "ELH": "Hopkins Fulfillment Service",
    "HCCA": "Harper Collins Publishers",
    "NY": "Ingram Publisher Services",
    "CONS": "Ingram Publisher Services",
    "CA": "Cambridge University Press",
    "MIO": "Longleaf Services, Inc",
}

smart_quotes_map = {"“": '"', "”": '"', "‘": "'", "’": "'"}

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")


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
                messagebox.showingo("Success", "Successfully loaded CSV file.")
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
            if data_element is not None:
                row_data.append(data_element.text)
            else:
                row_data.append(None)
        data.append(row_data)

    df = pd.DataFrame(data, columns=columns, index=None)
    excel_file_name = Path(file_path).stem + ".xlsx"
    excel_save_location = os.path.join(desktop_path, excel_file_name)
    df.to_excel(excel_save_location, index=False)
    print(df)
    return df


def delete_junk_cols(df):
    # function to delete unneeded columns
    df.drop(columns=cols_to_delete, inplace=True)
    df.columns = [
        "ISBN 10",
        "EAN",
        "Title",
        "Subtitle",
        "Title:Subtitle",
        "Author",
        "List Price",
        "Cost",
        "Edition",
        "PubDate",
        "Print Run",
        "Series",
        "Format Description",
        "Store Category",
        "BISAC Category Description",
        "Publisher Name",
        "Vendor",
    ]
    print(f"Columns {cols_to_delete} deleted.")
    return df


def move_start_article(text, prefixes):
    for prefix in prefixes:
        if isinstance(text, str) and text.startswith(prefix):
            return text[len(prefix) :].strip() + ", " + prefix
        return text


def item_name_number(row):
    if str(row["Subtitle"]) != "nan":
        row["item_name_number"] = (
            f"{row['Display_Name'] + '-' + row['Subtitle'] + '-' + str(row['UPC_Code'])}"
        )
        row["Long_Title"] = f"{row['Display_Name'].strip() + ': ' + row['Subtitle']}"
    else:
        row["item_name_number"] = f"{row['Display_Name'] + '-' + str(row['UPC_Code'])}"
        row["Long_Title"] = f"{row['Display_Name']}"
    return row


def assign_dept(row):
    if row["Class"].startswith("TRN"):
        row["Department"] = "Tradebooks : TR Nonfiction (TRN)"
        row["Income_Account"] = "325"
    elif row["Class"].startswith("TRF"):
        row["Department"] = "Tradebooks : TR Fiction & Lit (TRF)"
        row["Income_Account"] = "325"
    elif row["Class"].startswith("TRZ"):
        row["Department"] = "Tradebooks : TR Gifts (TRZ)"
        row["Income_Account"] = "1561"
    elif row["Class"].startswith("TRR"):
        row["Department"] = "Tradebooks : TR Reference (TRR)"
        row["Income_Account"] = "1549"
    return row


def rename_and_add_cols(df):
    current_cols = df.columns.tolist()
    df1 = pd.DataFrame(df)
    df2 = pd.DataFrame()
    first_row_idx = 0
    last_row_idx = len(df)
    print(current_cols)
    df1["PubDate"] = pd.to_datetime(df1["PubDate"], format="%m/%d/%Y")
    df1["Edition"] = df1["Edition"].fillna("N/A")

    df["ISBN 10"] = df.apply(
        lambda row: f"{row['EAN']}", axis=1
    )  # copies EAN/UPC to ISBN column

    df1.rename(
        columns={
            "ISBN 10": "ISBN",
            "EAN": "UPC_Code",
            "Title": "Display_Name",
            "Author": "Author",
            "List Price": "MSRP",
            "Cost": "Cost",
            "Edition": "Edition",
            "PubDate": "Publish_Date",
            "Print Run": "Print_Run",
            "Format Description": "Cover_Type",
            "Store Category": "Class",
            "BISAC Category Description": "BISAC",
            "Publisher Name": "Publisher",
            "Vendor": "Vendor",
        },
        inplace=True,
    )

    definite_articles = ["The ", "A ", "the ", "a "]
    df1["Display_Name"] = df1["Display_Name"].apply(
        lambda x: move_start_article(x, definite_articles)
    )

    new_cols = {
        "External_ID": "",
        "Long_Title": "",
        "Price_Level": "Base Price",
        "Department": "",
        "Income_Account": "",
        "Shelving_Category": "",
        "Web_Description": "",
        "Preferred_Location": "Main Campus Bookstore",
        "Webstore_Image_Name": "",
    }
    for index in range(first_row_idx, last_row_idx):
        for col, val in new_cols.items():
            df2.loc[index, col] = val

    alt_df = pd.merge(df1, df2, left_index=True, right_index=True)

    alt_df = alt_df.apply(item_name_number, axis=1)

    alt_df = alt_df.drop(columns=["Subtitle", "Title:Subtitle"])

    alt_df["External_ID"] = alt_df["UPC_Code"]

    alt_df["Vendor"] = alt_df["Vendor"].replace(vendor_dict)

    alt_df["Publish_Date"] = alt_df["Publish_Date"].dt.strftime("%m/%d/%y")

    alt_df["ISBN"] = alt_df["UPC_Code"]

    alt_df["Webstore_Image_Name"] = alt_df["UPC_Code"]

    col_order = [
        "External_ID",
        "ISBN",
        "UPC_Code",
        "Webstore_Image_Name",
        "item_name_number",
        "Display_Name",
        "Long_Title",
        "Author",
        "MSRP",
        "Price_Level",
        "Cost",
        "Edition",
        "Publish_Date",
        "Print_Run",
        "Cover_Type",
        "Class",
        "Department",
        "Income_Account",
        "BISAC",
        "Shelving_Category",
        "Publisher",
        "Vendor",
        "Web_Description",
        "Series",
        "Preferred_Location",
    ]
    alt_df = alt_df.apply(assign_dept, axis=1)
    print(alt_df)

    return alt_df[col_order]


def set_cover(row):
    if "Hardcover" in row["Cover_Type"]:
        row["Cover_Type"] = "Hardcover"
    elif "Other" in row["Cover_Type"]:
        row["Cover_Type"] = "Non-book item"
    else:
        return row
    return row


def get_images_and_desc(isbn):
    base_url = f"https://api2.isbndb.com/book/"
    headers = {"accept": "application/json", "Authorization": api_key}

    os.makedirs(dl_folder_home, exist_ok=True)
    book_cover = f"{isbn}.jpg"
    images_save_path = os.path.join(dl_folder_home, book_cover)

    try:
        url = f"{base_url}{isbn}"
        response = req.get(url, headers=headers, stream=True)
        print(headers)
        response.raise_for_status()
        book_data = response.json()
        if "synopsis" in book_data["book"] and "image_original" in book_data["book"]:
            synopsis = book_data["book"]["synopsis"]
            image = book_data["book"]["image_original"]
            image_link = req.get(image, stream=True)
            image_link.raise_for_status()
            with open(images_save_path, "wb") as file:
                for chunk in image_link.iter_content(chunk_size=8192):
                    file.write(chunk)
            return synopsis
        elif "synopsis" not in book_data["book"]:
            messagebox.showerror(
                "Error",
                f"No sypnosis available for {isbn}. Using title as Web Description.",
            )
            return book_data["book"]["title_long"]
        elif "image" not in book_data["image_original"]:
            messagebox.showerror("Error", f"No image found for {isbn}.")
            return
        time.sleep(3)
    except HTTPError as e:
        print(f"HTTP Error occurred: {e}")
        print(f"Status code: {e.response.status_code}")
        if e.response.status_code == 404:
            print(
                f"No result for ISBN {isbn}. Try at a later date to check if a record exists."
            )
        elif e.response.status_code == 429:
            print(f"Daily requests limit met. Please try tomorrow.")
    except req.exceptions.RequestException as e:
        messagebox.showerror("Error", f"API call failed for {isbn}: \n{e}")
        return None


def resize_covers(input_folder, output_folder, canvas_size):
    output_dir = Path(desktop_path) / output_folder

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for filename in os.listdir(input_folder):
        if filename.lower().endswith((".jpg", ".jpeg")):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_dir, filename)

            try:
                img = Image.open(input_path).convert("RGB")
            except Exception as e:
                print(f"Error processing {filename}: {e}")
                return

            canvas_width, canvas_height = canvas_size
            img_width, img_height = img.size

            canvas = Image.new("RGB", canvas_size, (255, 255, 255))

            if img_width > canvas_width or img_height > canvas_height:
                ratio_w = canvas_width / img_width
                ratio_h = canvas_height / img_height
                scale_factor = min(ratio_w, ratio_h)

                new_width = int(img_width * scale_factor)
                new_height = int(img_height * scale_factor)
                resized_img = img.resize(
                    (new_width, new_height), Image.Resampling.LANCZOS
                )

                x_offset = (canvas_width - new_width) // 2
                y_offset = (canvas_height - new_height) // 2
                canvas.paste(resized_img, (x_offset, y_offset))
            else:
                x_offset = (canvas_width - img_width) // 2
                y_offset = (canvas_height - img_height) // 2
                canvas.paste(img, (x_offset, y_offset))

            canvas.save(output_path)
            print(f"Processed {filename} saved to {output_path}")


def loading_popup(root, task_functions, title="Working", message="Please wait..."):
    popup = tk.Toplevel(root)
    popup.title(title)
    tk.Label(popup, text=message).pack(padx=20, pady=10)

    progress_bar = ttk.Progressbar(popup, mode="indeterminate", length=200)
    progress_bar.pack(padx=20, pady=10)
    progress_bar.start(10)

    def task_wrapper():
        change_po(*task_functions)
        root.after(1000, popup.destroy)

    worker_thread = threading.Thread(target=task_wrapper, daemon=True)
    worker_thread.start()


def change_po(
    load_file,
    delete_junk_cols,
    rename_and_add_cols,
    set_cover,
    get_images_and_desc,
    resize_covers,
):
    try:
        df = load_file(root)
        if df is None:
            return

        delete_junk_cols(df)
        print(df)
        df = rename_and_add_cols(df)
        print(df)
        df = df.apply(set_cover, axis=1)
        df["Web_Description"] = df["UPC_Code"].apply(get_images_and_desc)
        for smart, standard in smart_quotes_map.items():
            df["Web_Description"] = df["Web_Description"].str.replace(
                smart, standard, regex=False
            )
        resize_covers(dl_folder_home, f"{date_string}-resized_images", (600, 600))

        archive_name = os.path.join(desktop_path, f"{resized_images}-compressed")
        archive_format = "zip"
        shutil.make_archive(archive_name, archive_format)
        messagebox.showinfo(
            "Compressed",
            f"Image folder compressed for upload and saved to {dl_folder_home}.",
        )

        try:
            shutil.rmtree(dl_folder_home)
            print(f"\nSuccessfully deleted original image folder {dl_folder_home}.")
        except OSError as e:
            print(f"\nError deleting folder: {dl_folder_home}: {e}")

        save_directory = filedialog.askdirectory(
            parent=root, title="Select Save Location"
        )
        excel_name = f"Tradebooks_Export_{date_string}.xlsx"
        csv_name = f"Tradebooks_Export_{date_string}.csv"
        excel_save_path = os.path.join(save_directory, excel_name)
        csv_save_path = os.path.join(save_directory, csv_name)

        with pd.ExcelWriter(excel_save_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=excel_name, index=False)
            format_cols = [
                "External_ID",
                "ISBN",
                "UPC_Code",
                "Webstore_Image_Name",
            ]
            for col_name in format_cols:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in writer.sheets[excel_name].iter_rows(
                    min_col=col_idx, max_col=col_idx
                ):
                    for cell in row:
                        cell.number_format = "0"
        df.to_csv(csv_save_path, index=False)
        messagebox.showinfo(
            "Success",
            f"Processed Excel file saved to {excel_save_path}. \n\nProcessed CSV file saved to {csv_save_path}.",
        )
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Loading...")
    loading_popup(
        root,
        [
            load_file,
            delete_junk_cols,
            rename_and_add_cols,
            set_cover,
            get_images_and_desc,
            resize_covers,
        ],
    )

    root.withdraw()
    root.mainloop()
