#!/usr/bin/env python3

import os
import pandas as pd
import ipaddress
from typing import Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox


class ExcelHandler:
    def __init__(self, file_path: str, sheet_name: Optional[str] = None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df: Optional[pd.DataFrame] = None
        self.df_dict: dict[str, pd.DataFrame] = {}

    def read(self) -> pd.DataFrame:
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        xls = pd.ExcelFile(self.file_path)
        if self.sheet_name and self.sheet_name in xls.sheet_names:
            self.df = pd.read_excel(xls, sheet_name=self.sheet_name)
        else:
            self.sheet_name = xls.sheet_names[0]
            self.df = pd.read_excel(xls, sheet_name=self.sheet_name)

        self.df_dict.clear()
        self.df_dict["original"] = self.df.copy()
        return self.df

    def validate_ipv4_addresses(self, ip_column: str = "IP ADDRESS") -> None:
        if self.df is None:
            raise ValueError("Data not loaded.")
        self.df[ip_column] = self.df[ip_column].astype(str).str.strip()

        def is_valid_ipv4(ip: str) -> bool:
            try:
                ip_obj = ipaddress.IPv4Address(ip)
                return not (ip_obj.is_loopback or ip_obj.is_unspecified or ip_obj == ipaddress.IPv4Address("255.255.255.255"))
            except ipaddress.AddressValueError:
                return False

        is_valid = self.df[ip_column].apply(is_valid_ipv4)
        is_duplicated = self.df.duplicated(subset=[ip_column], keep=False)
        is_invalid = ~is_valid | is_duplicated

        self.df_dict["valid_ips"] = self.df[~is_invalid]
        self.df_dict["invalid_and_repeated_ips"] = self.df[is_invalid]

    def split_brand_match(self, brand_column: str = "BRAND") -> None:
        if self.df is None:
            raise ValueError("Data not loaded.")

        substrings = ['SLBS', '2L1T', '3L1T', '4L1T', '2L2T', '3L', '4L']
        pattern = '|'.join(substrings)

        self.df[brand_column] = self.df[brand_column].astype(str).str.strip()
        mask = self.df[brand_column].str.contains(pattern, case=False, na=False)

        self.df_dict["brand_matched"] = self.df[mask]
        self.df_dict["brand_unmatched"] = self.df[~mask]

    def export_results(self, output_file: str = "filtered_output.xlsx") -> None:
        if not self.df_dict:
            print("⚠️ No data to export.")
            return

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for sheet_name, df in self.df_dict.items():
                if df is not None and not df.empty:
                    clean_name = sheet_name[:31].replace('/', '_').replace('\\', '_')
                    df.to_excel(writer, sheet_name=clean_name, index=False)


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
