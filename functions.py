import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from data import cols_to_delete, vendor_dict, smart_quotes_map, trade_org
import os
from pathlib import Path
import requests as req
from requests.exceptions import HTTPError
from PIL import Image
from dotenv import load_dotenv
import time
import os
from file_handling import desktop_path
import shutil

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


def delete_junk_cols(df):
    # function to delete unneeded columns
    df = df.drop(columns=cols_to_delete)
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
    class_val = row["Class"]
    if pd.isna(class_val) or class_val is None:
        row["needs_manual_category"] = True
        return row
    if class_val.startswith("TRN"):
        row["Department"] = "Tradebooks : TR Nonfiction (TRN)"
        row["Income_Account"] = "325"
    elif class_val.startswith("TRF"):
        row["Department"] = "Tradebooks : TR Fiction & Lit (TRF)"
        row["Income_Account"] = "325"
    elif class_val.startswith("TRZ"):
        row["Department"] = "Tradebooks : TR Gifts (TRZ)"
        row["Income_Account"] = "1561"
    elif class_val.startswith("TRR"):
        row["Department"] = "Tradebooks : TR Reference (TRR)"
        row["Income_Account"] = "1549"
    else:
        row["Department"] = None
        row["Income_Account"] = None
        row["needs_manual_category"] = True
    return row


def update_row_with_selections(dept, cat, row_data, final_po, queue):
    try:
        income_account_map = {
            "Tradebooks : TR Nonfiction (TRN)": "325",
            "Tradebooks : TR Fiction & Lit (TRF)": "325",
            "Tradebooks : TR Gifts (TRZ)": "1561",
            "Tradebooks : TR Reference (TRR)": "1549",
        }

        matching_rows = final_po[
            final_po["item_name_number"] == row_data["item_name_number"]
        ]
        if matching_rows.empty:
            queue.put((False, "Row not found in dataframe"))
            return

        row_idx = matching_rows.index[0]

        final_po.loc[row_idx, "Department"] = dept
        final_po.loc[row_idx, "Class"] = cat
        final_po.loc[row_idx, "Income_Account"] = income_account_map.get(dept, "")
        final_po.loc[row_idx, "needs_manual_category"] = False

        queue.put((True, None))
    except Exception as e:
        queue.put((False, str(e)))


def rename_and_add_cols(df, po_number):
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
            "Units": "Qty",
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
        "PO_Num": po_number,
        "External_ID": "",
        "Long_Title": "",
        "Price_Level": "Base Price",
        "Department": "",
        "Income_Account": "",
        "Shelving_Category": "",
        "Web_Description": "",
        "Preferred_Location": "Main Campus Bookstore",
        "Webstore_Image_Name": "",
        "needs_manual_category": False,
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
        "PO_Num",
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
        "Qty",
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
        "needs_manual_category",
    ]
    alt_df = alt_df.apply(assign_dept, axis=1)
    print(alt_df)

    return alt_df[col_order]


def set_cover(row):
    cover_type = row["Cover_Type"]
    if pd.isna(cover_type):
        row["Cover_Type"] = ""
        return row

    cover_type = str(cover_type).strip()

    if "Hardcover" in cover_type:
        row["Cover_Type"] = "Hardcover"
    elif "Other" in cover_type:
        row["Cover_Type"] = "Non-book item"
    else:
        return row
    return row


def get_images_and_desc(isbn, po_number, dl_folder_home):
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


def resize_covers(input_folder, po_number, canvas_size=(600, 600)):
    output_folder = f"{po_number}-resized_images"
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

    compress_images(po_number)

    try:
        shutil.rmtree(input_folder)
        print(f"\nSuccessfully deleted original image folder {input_folder}.")
    except OSError as e:
        print(f"\nError deleting folder: {input_folder}: {e}")


def compress_images(po_number):
    resized_folder_name = f"{po_number}-resized_images"
    resized_folder_path = os.path.join(desktop_path, resized_folder_name)
    archive_name = os.path.join(desktop_path, f"{po_number}-images-compressed")
    archive_format = "zip"

    shutil.make_archive(
        archive_name,
        archive_format,
        root_dir=desktop_path,
        base_dir=resized_folder_name,
    )

    messagebox.showinfo(
        "Compressed",
        f"Image folder compressed for upload and saved to {archive_name}.",
    )


def change_po(df, po_number):
    try:
        dl_folder_name = f"book_covers-{po_number}"
        dl_folder_home = os.path.join(desktop_path, dl_folder_name)

        df = delete_junk_cols(df)
        df = rename_and_add_cols(df, po_number)
        df = df.apply(set_cover, axis=1)

        df["Web_Description"] = df["UPC_Code"].apply(
            lambda isbn: get_images_and_desc(isbn, po_number, dl_folder_home)
        )

        for smart, standard in smart_quotes_map.items():
            df["Web_Description"] = df["Web_Description"].str.replace(
                smart, standard, regex=False
            )

        resize_covers(dl_folder_home, po_number)

        return df

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        return None
