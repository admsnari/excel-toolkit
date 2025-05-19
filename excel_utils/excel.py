import os
import pandas as pd
import ipaddress
from typing import Optional



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
