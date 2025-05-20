#!/usr/bin/env python3

import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Listbox
from excel_utils.excel import ExcelHandler

def run_app():
    def reset_app():
        file_path.set("")
        sheet_name_input.delete(0, END)
        ip_check_var.set(False)
        brand_check_var.set(False)
        analyze_wizard_var.set(False)
        status_label.config(text="")

    def select_file():
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title="Select Excel File"
        )
        if path:
            file_path.set(path)
            status_label.config(text=f"Selected: {os.path.basename(path)}")

    def process_file():
        try:
            fp = file_path.get()
            sheet = sheet_name_input.get().strip() or None
            if not fp:
                messagebox.showwarning("Warning", "No file selected!")
                return

            handler = ExcelHandler(fp, sheet_name=sheet)
            df = handler.read()

            ip_col = "IP ADDRESS"
            brand_col = "BRAND"

            if analyze_wizard_var.get() and brand_col not in df.columns and ip_col not in df.columns:
                show_column_names("Missing IP ADDRESS and BRAND Column", df.columns)
                ip_col_custom, brand_col_custom = get_custom_ip_and_brand_name("Wizard: Enter Custom Column Names")
                if not brand_col_custom or not ip_col_custom:
                    messagebox.showerror("Error", "Both Brand and IP columns are required.")
                    return
                brand_col = brand_col_custom
                ip_col = ip_col_custom

            if ip_check_var.get() and ip_col not in df.columns:
                show_column_names("Missing IP ADDRESS Column", df.columns)
                ip_col_custom = get_custom_column_name("Missing IP Column", "Enter custom IP column name:")
                if not ip_col_custom:
                    messagebox.showerror("Error", "No IP column name provided.")
                    return
                ip_col = ip_col_custom

            if brand_check_var.get() and brand_col not in df.columns:
                show_column_names("Missing BRAND Column", df.columns)
                brand_col_custom = get_custom_column_name("Missing Brand Column", "Enter custom Brand column name:")
                if not brand_col_custom:
                    messagebox.showerror("Error", "No Brand column name provided.")
                    return
                brand_col = brand_col_custom


            if ip_check_var.get():
                handler.validate_ipv4_addresses(ip_column=ip_col)
            if brand_check_var.get():
                handler.split_brand_match(brand_column=brand_col)
            if analyze_wizard_var.get():
                handler.analyze_wizard(ip_column=ip_col, brand_column=brand_col)


            output = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Save Output File"
            )
            if not output:
                return

            handler.export_results(output_path=output)
            messagebox.showinfo("Success", f"Results saved to:\n{output}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def show_help():
        messagebox.showinfo("Help & Instructions", """\
    ✔️ Brand filter will match the following patterns: SLBS, 2L1T, 3L1T, 4L1T, 2L2T, 3L, 4L
    ✔️ 'Validate IP Addresses' checks for valid IPv4 format and removes duplicates.
    ✔️ 'Filter by Brand' filters rows that contain specific brand patterns.
    ✔️ The default IP and Brand Column Name: "IP ADDRESS" "BRAND".
    ✔️ Use the 'ADMS Excel Process Wizard 🚀' for an automated one-click process.
    📌 Tip: You can provide custom column names if your Excel columns are different.""")
        
    # show the user excel columns if the user file is not have the defult columns name
    def show_column_names(title, columns):
        col_win = tb.Toplevel()
        col_win.title(title)
        col_win.geometry("800x700")
        col_win.grab_set()  # Modal

        tb.Label(col_win, text="Available columns in the Excel file:").pack(pady=10)

        listbox = Listbox(col_win, height=15, width=50)
        listbox.pack(pady=5)

        for col in columns:
            listbox.insert("end", col)

        tb.Button(col_win, text="OK", command=col_win.destroy, bootstyle="info").pack(pady=10)
        root.wait_window(col_win)


    # get one custom column name.
    def get_custom_column_name(title, label_text):
        input_win = tb.Toplevel()
        input_win.title(title)
        input_win.geometry("400x150")
        input_win.grab_set()  # Make it modal

        tb.Label(input_win, text=label_text).pack(pady=10)
        input_entry = tb.Entry(input_win, width=30)
        input_entry.pack(pady=5)

        result = {"value": None}

        def submit():
            result["value"] = input_entry.get().strip()
            input_win.destroy()

        tb.Button(input_win, text="Submit", command=submit, bootstyle="success").pack(pady=10)
        root.wait_window(input_win)
        return result["value"]

    # get custom ip and brand column name
    def get_custom_ip_and_brand_name(title):
        input_win = tb.Toplevel()
        input_win.title(title)
        input_win.geometry("400x200")
        input_win.grab_set()  # Make it modal

        # IP Address Label & Entry
        tb.Label(input_win, text="Enter custom IP column name:").pack(pady=5)
        ip_entry = tb.Entry(input_win, width=30)
        ip_entry.pack(pady=5)

        # Brand Label & Entry
        tb.Label(input_win, text="Enter custom Brand column name:").pack(pady=5)
        brand_entry = tb.Entry(input_win, width=30)
        brand_entry.pack(pady=5)

        result = {"ip": None, "brand": None}

        def submit():
            result["ip"] = ip_entry.get().strip()
            result["brand"] = brand_entry.get().strip()
            input_win.destroy()

        tb.Button(input_win, text="Submit", command=submit, bootstyle="success").pack(pady=10)
        root.wait_window(input_win)

        return result["ip"], result["brand"]



    # Set up themed window
    root = tb.Window(themename="flatly")
    root.title("Excel IP & Brand Filter - NARI 🚀")
    root.geometry("1500x820")
    root.resizable(True, True)

    file_path = tb.StringVar()
    ip_check_var = tb.BooleanVar()
    brand_check_var = tb.BooleanVar()
    analyze_wizard_var = tb.BooleanVar()

    tb.Label(root, text="Excel Sheet Processor (Team: NARI)", font=("Segoe UI", 16, "bold")).pack(pady=10)
    # tb.Label(root, text="Brand filter matches: SLBS, 2L1T, 3L1T, 4L1T, 2L2T, 3L, 4L", style="info").pack()
    # tb.Label(root, text="If you are new to this app, please use the ADMS Excel Process Wizard!", style="info").pack()

    tb.Button(root, text="📂 Browse Excel File", command=select_file, style=PRIMARY).pack(pady=10)
    status_label = tb.Label(root, text="", style="secondary")
    status_label.pack()

    tb.Label(root, text="Sheet name (leave blank for default):").pack(pady=5)
    sheet_name_input = tb.Entry(root, width=30)
    sheet_name_input.pack()

    # Create a frame to hold the checkbuttons
    check_frame = tb.Frame(root, padding=10)
    check_frame.pack(pady=20)

    # Add checkbuttons to the frame with left alignment
    tb.Checkbutton(check_frame, text="Validate IP Addresses", variable=ip_check_var, bootstyle="success").pack(pady=5, anchor="w", fill="x")
    tb.Checkbutton(check_frame, text="Filter by Brand", variable=brand_check_var, bootstyle="success").pack(pady=5, anchor="w", fill="x")
    tb.Checkbutton(check_frame, text="ADMS EXCEL PROCESS WIZARD 🚀", variable=analyze_wizard_var, bootstyle="success").pack(pady=5, anchor="w", fill="x")


    # Button Frame to center all buttons
    button_frame = tb.Frame(root)
    button_frame.pack(pady=20)

    # Common width for all buttons
    button_width = 25

    tb.Button(button_frame, text="🚀 Process & Export", command=process_file, bootstyle=SUCCESS, width=button_width).pack(pady=5)
    tb.Button(button_frame, text="🔄 Reset App", command=reset_app, bootstyle=DANGER, width=button_width).pack(pady=5)
    tb.Button(button_frame, text="❓ Help & Instructions", command=show_help, bootstyle=(INFO, OUTLINE), width=button_width).pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    run_app()
