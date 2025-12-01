import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import threading
from dotenv import load_dotenv
import requests as req
from PIL import Image
import shutil

load_dotenv(dotenv_path="./dist/ISBNdb_API_Key.env")
api_key = os.getenv("API_KEY")
print(api_key)

cols_to_delete = ['Purchase Order', 'Item #', 'Author 2', 'Units', 'Discount', 'Discount Code', 'Format', 'Secondary Store Category', 'BISAC Category Code', 'Publisher Code', 'SAN', 'Markup Tags', 'UPC', 'Number of Pages', 'Dimensions', 'Links']

cols_dtypes = {
    'ISBN 10': str,
    'EAN': str,
    'Title': str,
    'Subtitle': str,
    'List Price': float,
    'Cost': float,
    'Edition': str,
    'Print Run': str
}

vendor_dict = {
    'Penguin Random House': 'Penguin Random House LLC',
    'Simon & Schuster': 'Simon & Schuster',
    'MPS': 'MPS',
    'HC': 'Harper Collins Publishers',
    'IN': 'Ingram Publisher Services',
    'TW': 'Hachette Book Group USA',
    'NO': 'W W Norton & Company Inc',
    '394': 'Penguin Random House LLC',
    'WI': 'John Wiley & Songs Inc',
    'JHU': 'Indiana University - Slavica Publishers',
    'ABRA': 'Hachette Book Group USA',
    'CM': 'Ingram Publisher Services',
    'LCQ': 'Harper Collins Publishers',
    'HARB': 'W W Norton & Company Inc',
    'DMMQ': 'Longleaf Services, Inc',
    'CUNW': 'Ingram Publisher Services',
    'BKVR': 'Taylor & Francis Group, LLC',
    'REDB': 'Red Wheel Weiser, LLC',
    'UOCP': 'Ingram Publisher Services',
    'ELH': 'Hopkins Fulfillment Service',
    'HCCA': 'Harper Collins Publishers',
    'NY': 'Ingram Publisher Services',
    'CONS': 'Ingram Publisher Services',
    'CA': 'Cambridge University Press',
    'MIO': 'Longleaf Services, Inc'
}

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")

def load_file(root):
    file_path = filedialog.askopenfilename(
        title="Select input file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting")
    
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, dtype=cols_dtypes, header=0)
            messagebox.showinfo("Success", "Successfully loaded Excel file.")
            return df
        elif file_path.endswith('.xls'):
            try:
                df = pd.read_xls(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded Excel file.")
                return df
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load. Please try opening the file and saving as .xls or .xlsx, then run this script again.")
                return None
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=0)
            messagebox.showinfo("Success", "Successfully loaded CSV file.")
            return df
        else:
            messagebox.showerror("Error", "Unsupported file format. Please provide a valid Excel file (.xlsx).")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading file: {e}")
        return None

def delete_junk_cols(df):
    # function to delete unneeded columns
    df.drop(columns=cols_to_delete, inplace=True)
    df.columns = ['ISBN 10', 'EAN', 'Title', 'Subtitle', 'Title:Subtitle', 'Author', 'List Price', 'Cost', 'Edition', 'PubDate', 'Print Run', 'Series', 'Format Description', 'Store Category', 'BISAC Category Description', 'Publisher Name', 'Vendor']
    print(f"Columns {cols_to_delete} deleted.")
    return df

def move_start_article(text, prefixes):
    for prefix in prefixes:
        if isinstance(text, str) and text.startswith(prefix):
            return text[len(prefix):].strip() + ', ' + prefix
        return text

def item_name_number(row):
    if str(row['Subtitle']) != 'nan':
        row['item_name_number'] = f'{row['displayname'] + '-' + row['Subtitle'] + '-' + str(row['upccode'])}'
        row['custitem_nsts_csic_long_title'] = f'{row['displayname'].strip() + ': ' + row['Subtitle']}'
    else:
        row['item_name_number'] = f'{row['displayname'] + '-' + str(row['upccode'])}'
        row['custitem_nsts_csic_long_title'] = f'{row['displayname']}'
    return row

def assign_dept(row):
    if row['class'].startswith('TRN'):
        row['department'] = 'Tradebooks : TR Nonfiction (TRN)'
        row['incomeaccount'] = '325'
    elif row['class'].startswith('TRF'):
        row['department'] = 'Tradebooks : TR Fiction & Lit (TRF)'
        row['incomeaccount'] = '325'
    elif row['class'].startswith('TRZ'):
        row['department'] = 'Tradebooks : TR Gifts (TRZ)'
        row['incomeaccount'] = '1561'
    elif row['class'].startswith('TRR'):
        row['department'] = 'Tradebooks : TR Reference (TRR)'
        row['incomeaccount'] = '1549'
    return row

def rename_and_add_cols(df):
    current_cols = df.columns.tolist()
    df1 = pd.DataFrame(df)
    df2 = pd.DataFrame()
    first_row_idx = 0
    last_row_idx = len(df)
    print(current_cols)
    df1['PubDate'] = pd.to_datetime(df1['PubDate'], format='%m/%d/%Y')
    df1['Edition'] = df1['Edition'].fillna('N/A')

    df1.rename(columns={'ISBN 10': 'custitem_nsts_csic_isbn', 'EAN': 'upccode', 'Title': 'displayname', 'Author': 'custitem_nsts_csic_author', 'List Price': 'listprice', 'Cost': 'costestimate', 'Edition': 'custitem_nsts_csic_edition', 'PubDate': 'custitem_nsts_csic_pub_date', 'Print Run': 'custitem_nsts_csic_printing', 'Format Description': 'custitem_nsts_csic_cover_type', 'Store Category': 'class', 'BISAC Category Description': 'custitemcustitem_bisac', 'Publisher Name': 'custitem_nsts_csic_imprint_pub', 'Vendor': 'vendor'}, inplace=True)

    definite_articles = ['The ', 'A ', 'the ', 'a ']
    df1['displayname'] = df1['displayname'].apply(lambda x: move_start_article(x, definite_articles))

    new_cols = {'externalid': '', 'custitem_nsts_csic_long_title' : '', 'pricelevel': 'Base Price', 'department': '', 'incomeaccount': '', 'shelving_category': '', 'Web Description': '', 'Preferred_Location': 'Main Campus Bookstore', 'Webstore Image Name': ''}
    for index in range(first_row_idx, last_row_idx):
        for col, val in new_cols.items():
            df2.loc[index, col] = val
            
    alt_df = pd.merge(df1, df2, left_index=True, right_index=True)

    alt_df = alt_df.apply(item_name_number, axis=1)

    alt_df = alt_df.drop(columns=['Subtitle', 'Title:Subtitle'])

    alt_df['externalid'] = alt_df['custitem_nsts_csic_isbn']

    alt_df['vendor'] = alt_df['vendor'].replace(vendor_dict)

    alt_df['custitem_nsts_csic_pub_date'] = alt_df['custitem_nsts_csic_pub_date'].dt.strftime('%m/%d/%y')

    alt_df['Webstore Image Name'] = alt_df['upccode']

    col_order = ['externalid', 'custitem_nsts_csic_isbn', 'upccode', 'Webstore Image Name', 'item_name_number', 'displayname', 'custitem_nsts_csic_long_title', 'custitem_nsts_csic_author', 'listprice', 'pricelevel', 'costestimate', 'custitem_nsts_csic_edition', 'custitem_nsts_csic_pub_date', 'custitem_nsts_csic_printing', 'custitem_nsts_csic_cover_type', 'class', 'department', 'incomeaccount', 'custitemcustitem_bisac', 'shelving_category', 'custitem_nsts_csic_imprint_pub', 'vendor', 'Web Description', 'Series', 'Preferred_Location']
    alt_df = alt_df.apply(assign_dept, axis=1)
    print(alt_df)

    return alt_df[col_order]

def set_cover(row):
    if 'Hardcover' in row['custitem_nsts_csic_cover_type']:
        row['custitem_nsts_csic_cover_type'] = 'Hardcover'
    elif 'Other' in row['custitem_nsts_csic_cover_type']:
        row['custitem_nsts_csic_cover_type'] = "Non-book item"
    else:
        return row
    return row

def get_images_and_desc(isbn):
    base_url = f'https://api2.isbndb.com/book/'
    headers = {
    'accept': 'application/json',
    'Authorization': api_key
    }
    current_directory = os.getcwd()
    dl_folder_name = f'book_covers-{date_string}'
    dl_folder_home = os.path.join(current_directory, dl_folder_name)
    os.makedirs(dl_folder_home, exist_ok=True)
    book_cover = f'{isbn}.jpg'
    images_save_path = os.path.join(dl_folder_home, book_cover)

    try:
        url = f'{base_url}{isbn}'
        response = req.get(url, headers=headers, stream=True)
        print(headers)
        response.raise_for_status()
        book_data = response.json()
        if 'synopsis' in book_data['book'] and 'image_original' in book_data['book']:
            synopsis = book_data['book']['synopsis']
            image = book_data['book']['image_original']
            image_link = req.get(image, stream=True)
            image_link.raise_for_status()
            with open(images_save_path, 'wb') as file:
                for chunk in image_link.iter_content(chunk_size=8192):
                    file.write(chunk)
            resize_covers(dl_folder_home, f"{date_string}-resized_images", (600, 600))
            try:
                shutil.rmtree(dl_folder_home)
                print(f"\nSuccessfully deleted original image folder {dl_folder_home}.")
            except OSError as e:
                print(f"\nError deleting folder: {dl_folder_home}: {e}")
            return synopsis
        elif 'synopsis' not in book_data['book']:
            messagebox.showerror('Error', f'No sypnosis available for {isbn}. Using title as Web Description.')
            return book_data['book']['title_long']
        elif 'image' not in book_data['image_original']:
            messagebox.showerror('Error', f'No image found for {isbn}. Skipping.')
            return
    except req.exceptions.RequestException as e:
        messagebox.showerror("Error", f"API call failed for {isbn}: {e}")
        return None

def resize_covers(input_folder, output_folder, canvas_size):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(('.jpg', '.jpeg')):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)

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
                resized_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

                x_offset = (canvas_width - new_width) // 2
                y_offset = (canvas_height - new_height) // 2
                canvas.paste(resized_img, (x_offset, y_offset))
            else:
                x_offset = (canvas_width - img_width) // 2
                y_offset = (canvas_height - img_height) // 2
                canvas.paste(img, (x_offset, y_offset))
            
            canvas.save(output_path)
            print(f"Processed {filename} saved to {output_path}")
    
    try:
        shutil.rmtree(input_folder)
        print(f"\nSuccessfully deleted original image folder {input_folder}.")
    except OSError as e:
        print(f"\nError deleting folder: {input_folder}: {e}")
    

def loading_popup(root, task_functions, title="Working", message="Please wait..."):
    popup = tk.Toplevel(root)
    popup.title(title)
    tk.Label(popup, text=message).pack(padx=20, pady=10)

    progress_bar = ttk.Progressbar(popup, mode='indeterminate', length=200)
    progress_bar.pack(padx=20, pady=10)
    progress_bar.start(10)

    def task_wrapper():
        change_po(*task_functions)
        root.after(1000, popup.destroy)
    
    worker_thread = threading.Thread(target=task_wrapper, daemon=True)
    worker_thread.start()

def change_po(load_file, delete_junk_cols, rename_and_add_cols, set_cover, get_images_and_desc):
    try:
        df = load_file(root)
        if df is None:
            return
        
        delete_junk_cols(df)
        print(df)
        df = rename_and_add_cols(df)
        print(df)
        df = df.apply(set_cover, axis=1)
        df['Web Description'] = df['upccode'].apply(get_images_and_desc)

        save_directory = filedialog.askdirectory(parent=root, title="Select Save Location")
        default_name = f'Tradebooks_Export_{date_string}.xlsx'
        full_save_path = os.path.join(save_directory, default_name)

        with pd.ExcelWriter(full_save_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=default_name, index=False)
            format_cols = ['externalid', 'custitem_nsts_csic_isbn', 'upccode']
            for col_name in format_cols:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in writer.sheets[default_name].iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '0'
        messagebox.showinfo("Success", f"Processed file saved to {full_save_path}")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Loading...")
    loading_popup(root, [load_file, delete_junk_cols, rename_and_add_cols, set_cover, get_images_and_desc])

    root.withdraw()
    root.mainloop()
