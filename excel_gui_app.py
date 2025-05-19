#!/usr/bin/env python3

from excel_utils.excel import ExcelHandler
import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox


def run_app():
    def reset_app():
        file_path.set("")
        sheet_name_input.delete(0, END)
        ip_col_entry.delete(0, END)
        brand_col_entry.delete(0, END)
        ip_col_frame.pack_forget()
        brand_col_frame.pack_forget()
        ip_check_var.set(False)
        brand_check_var.set(False)
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

            if ip_check_var.get() and ip_col not in df.columns:
                ip_col_frame.pack(pady=5)
                ip_col_custom = ip_col_entry.get().strip()
                if not ip_col_custom:
                    messagebox.showerror("Error", "Please provide IP column name.")
                    return
                ip_col = ip_col_custom

            if brand_check_var.get() and brand_col not in df.columns:
                brand_col_frame.pack(pady=5)
                brand_col_custom = brand_col_entry.get().strip()
                if not brand_col_custom:
                    messagebox.showerror("Error", "Please provide Brand column name.")
                    return
                brand_col = brand_col_custom

            if ip_check_var.get():
                handler.validate_ipv4_addresses(ip_column=ip_col)
            if brand_check_var.get():
                handler.split_brand_match(brand_column=brand_col)

            output = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Save Output File"
            )
            if not output:
                return

            handler.export_results(output_file=output)
            messagebox.showinfo("Success", f"Results saved to:\n{output}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Set up themed window
    root = tb.Window(themename="flatly")
    root.title("Excel IP & Brand Filter - NARI 🚀")
    root.geometry("1500x820")
    root.resizable(True, True)

    file_path = tb.StringVar()
    ip_check_var = tb.BooleanVar()
    brand_check_var = tb.BooleanVar()

    tb.Label(root, text="Excel Sheet Processor (Team: NARI)", font=("Segoe UI", 16, "bold")).pack(pady=10)
    tb.Label(root, text="Brand filter matches: SLBS, 2L1T, 3L1T, 4L1T, 2L2T, 3L, 4L", bootstyle="info").pack()

    tb.Button(root, text="📂 Browse Excel File", command=select_file, bootstyle=PRIMARY).pack(pady=10)
    status_label = tb.Label(root, text="", bootstyle="secondary")
    status_label.pack()

    tb.Label(root, text="Sheet name (leave blank for default):").pack(pady=5)
    sheet_name_input = tb.Entry(root, width=30)
    sheet_name_input.pack()

    tb.Checkbutton(root, text="Validate IP Addresses", variable=ip_check_var, bootstyle="success").pack(pady=5)
    tb.Checkbutton(root, text="Filter by Brand", variable=brand_check_var, bootstyle="warning").pack(pady=5)

    ip_col_frame = tb.Frame(root)
    tb.Label(ip_col_frame, text="Custom IP Column Name:").pack(side=LEFT)
    ip_col_entry = tb.Entry(ip_col_frame, width=20)
    ip_col_entry.pack(side=LEFT)

    brand_col_frame = tb.Frame(root)
    tb.Label(brand_col_frame, text="Custom Brand Column Name:").pack(side=LEFT)
    brand_col_entry = tb.Entry(brand_col_frame, width=20)
    brand_col_entry.pack(side=LEFT)

    tb.Button(root, text="🚀 Process & Export", command=process_file, bootstyle=SUCCESS, width=20).pack(pady=10)
    tb.Button(root, text="🔄 Reset App", command=reset_app, bootstyle=DANGER, width=20).pack()

    root.mainloop()


if __name__ == "__main__":
    run_app()
