import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from functions import *
from file_handling import *
import threading
import queue

final_po = None
result_queue = queue.Queue()
worker_thread = None
popup = None
po_number = None


def thread_wrapper(loaded_file, po_number, result_queue):
    global final_po
    try:
        final_po = change_po(loaded_file, po_number)
        if final_po is None:
            result_queue.put((False, "Processing failed"))
        else:
            result_queue.put((True, final_po))
    except Exception as e:
        result_queue.put((False, str(e)))


def on_select_click():
    global worker_thread, popup, po_number

    po_number = prompt_for_po_number(root)
    if po_number is None:
        return

    loaded_file = load_file(root)
    if loaded_file is None:
        return

    content.destroy()

    popup = tk.Toplevel(root)
    popup.grid()
    popup.title("Working")
    ttk.Label(popup, text="Please wait...").grid(
        column=0, row=0, columnspan=5, padx=10, pady=10
    )

    progress_bar = ttk.Progressbar(popup, mode="indeterminate", length=200)
    progress_bar.grid(column=0, row=2, columnspan=5, padx=10, pady=5)
    progress_bar.start()

    worker_thread = threading.Thread(
        target=thread_wrapper, args=(loaded_file, po_number, result_queue), daemon=True
    )
    worker_thread.start()

    check_thread_status(root)


def prompt_for_po_number(parent_window):
    po_number_dialog = tk.Toplevel(parent_window)
    po_number_dialog.title("Enter PO Number")
    po_number_dialog.focus_set()
    po_number_dialog.grab_set()

    ttk.Label(po_number_dialog, text="Enter PO Number:").grid(
        column=1, row=0, padx=10, pady=10
    )

    po_entry = ttk.Entry(po_number_dialog)
    po_entry.grid(column=1, row=1, padx=10, pady=10)
    po_entry.focus_set()

    result = []

    def submit_po(event=None):
        po_val = po_entry.get().strip()
        if not po_val:
            messagebox.showerror("Error", "PO Number cannot be empty")
            return
        result.append(po_val)
        po_number_dialog.destroy()

    submit_button = ttk.Button(
        po_number_dialog, text="Submit", command=submit_po, default="active"
    )
    submit_button.grid(column=1, row=2, pady=10)

    po_entry.bind("<Return>", submit_po)
    po_entry.bind("<KP_Enter>", submit_po)

    parent_window.wait_window(po_number_dialog)
    return result[0] if result else None


def check_thread_status(root):
    global worker_thread, popup, final_po

    if worker_thread and worker_thread.is_alive():
        root.after(100, lambda: check_thread_status(root))
        return

    if popup and popup.winfo_exists():
        popup.destroy()

    if final_po is not None:
        problem_rows = final_po[final_po["needs_manual_category"] == True]
        if not problem_rows.empty:
            first_problem = problem_rows.iloc[0]
            select_dept_cat(root, first_problem)
            return

    on_success()


def select_dept_cat(root, row_data):
    category_win = tk.Toplevel(root)
    category_win.title("Category Error")
    category_win.geometry("400x400+500+300")

    category_win.grid_columnconfigure(0, weight=1)
    category_win.grid_columnconfigure(4, weight=1)

    ttk.Label(
        category_win,
        text=f"Category not in NetSuite. Select correct Department and Category.",
        justify="center",
        width=40,
        wraplength=380,
    ).grid(column=0, row=1, columnspan=5, padx=10, pady=10)

    current_lbl = ttk.Label(
        category_win,
        text=f"Current Category: {row_data['Class']}",
        justify="center",
        anchor="center",
        width=40,
        wraplength=380,
    )
    current_lbl.grid(column=0, row=2, columnspan=5, padx=10, pady=10)

    book_title = ttk.Label(
        category_win,
        text=f"Title: {row_data['Display_Name']}",
        justify="center",
        anchor="center",
        width=40,
        wraplength=380,
    )
    book_title.grid(column=0, row=3, columnspan=5, padx=10, pady=10)

    book_format = ttk.Label(
        category_win,
        text=f"Format: {row_data['Cover_Type']}",
        justify="center",
        anchor="center",
        width=40,
        wraplength=380,
    )
    book_format.grid(column=0, row=4, columnspan=5, padx=10, pady=10)

    book_bisac = ttk.Label(
        category_win,
        text=f"BISAC: {row_data['BISAC']}",
        justify="center",
        anchor="center",
        width=40,
        wraplength=380,
    )
    book_bisac.grid(column=0, row=5, columnspan=5, padx=10, pady=10)

    dept = tk.StringVar(category_win)
    cat = tk.StringVar(category_win)

    dept_list = list(trade_org.keys())
    dropdown_dept = ttk.Combobox(
        category_win, values=dept_list, textvariable=dept, state="readonly", width=30
    )
    dropdown_dept.grid(column=0, row=6, columnspan=5, padx=10, pady=10)

    def on_dept_selected(event):
        selected_dept = dept.get()
        cat.set("")

        if selected_dept in trade_org:
            cat_dropdown["values"] = trade_org[selected_dept]
        else:
            cat_dropdown["values"] = []

    dropdown_dept.bind("<<ComboboxSelected>>", on_dept_selected)

    cat_dropdown = ttk.Combobox(
        category_win, textvariable=cat, state="readonly", width=30
    )
    cat_dropdown.grid(column=0, row=7, columnspan=5, padx=10, pady=10)

    def validate_and_submit(event=None):
        selected_dept = dept.get()
        selected_cat = cat.get()

        if not selected_dept:
            messagebox.showerror("Error", "Please select a Department")
            return

        if not selected_cat:
            messagebox.showerror("Error", "Please select a Category")
            return

        pass_selections(root, category_win, selected_dept, selected_cat, row_data)

    submit_cat = ttk.Button(
        category_win, text="Submit", command=validate_and_submit, default="active"
    )
    submit_cat.grid(column=0, row=8, columnspan=5, padx=10, pady=10)

    submit_cat.bind("<Return>", validate_and_submit)
    submit_cat.bind("<KP_Enter>", validate_and_submit)


def pass_selections(root, category_win, selected_dept, selected_cat, row_data):
    global worker_thread, popup
    category_win.destroy()

    popup = tk.Toplevel(root)
    popup.title("Working")
    ttk.Label(popup, text="Setting department and category...").grid(
        column=0, row=0, padx=10, pady=10
    )

    worker_thread = threading.Thread(
        target=update_row_with_selections,
        args=(selected_dept, selected_cat, row_data, final_po, result_queue),
        daemon=True,
    )
    worker_thread.start()

    check_thread_status(root)


def on_success():
    save_directory = filedialog.askdirectory(title="Select Save Location")

    if not save_directory:
        return

    file_name = f"Tradebooks_Export_{po_number}"
    save_path = os.path.join(save_directory, file_name)

    success, message = save_and_export(final_po, save_path, po_number)

    if success:
        messagebox.showinfo("Success", message)
    else:
        messagebox.showinfo("Error", f"Failed to save PO: {message}.")

    root.destroy()


root = tk.Tk()
root.title("Tradebook PO Prep")
root.geometry("250x250")

content = ttk.Frame(root)
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

load_btn = ttk.Button(
    content, text="Select PO file", command=on_select_click, default="active"
)

load_btn.bind("<Return>", on_select_click)
load_btn.bind("<KP_Enter>", on_select_click)

content.grid(column=0, row=0, columnspan=5)
load_btn.grid(column=0, row=0)

root.mainloop()
