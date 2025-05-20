import os
import re
import ipaddress
import pandas as pd
from typing import Optional


class ExcelHandler:
    """
    A class to handle Excel file operations including reading, validating IP addresses,
    brand pattern matching, and exporting analyzed results.
    """
    def __init__(self, file_path: str, sheet_name: Optional[str] = None):
        """
        Initialize ExcelHandler.

        :param file_path: Path to the Excel file.
        :param sheet_name: Optional sheet name to read.
        """
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df: Optional[pd.DataFrame] = None
        self.df_dict: dict[str, pd.DataFrame] = {}

    def read(self) -> pd.DataFrame:
        """
        Read the Excel file and load the specified or default sheet.

        :return: Loaded DataFrame.
        :raises FileNotFoundError: If the file does not exist.
        """
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
        """
        Validate IPv4 addresses in the specified column. Invalid or duplicate IPs are separated.

        :param ip_column: Name of the column containing IP addresses.
        :raises ValueError: If DataFrame is not loaded.
        """
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
        """
        Split the DataFrame into matched and unmatched brand patterns.

        :param brand_column: Name of the column containing brand identifiers.
        :raises ValueError: If DataFrame is not loaded.
        """
        if self.df is None:
            raise ValueError("Data not loaded.")

        substrings = ['SLBS', '2L1T', '3L1T', '4L1T', '2L2T', '3L', '4L']
        pattern = '|'.join(substrings)

        self.df[brand_column] = self.df[brand_column].astype(str).str.strip()
        mask = self.df[brand_column].str.contains(pattern, case=False, na=False)

        self.df_dict["brand_matched"] = self.df[mask]
        self.df_dict["brand_unmatched"] = self.df[~mask]

    def analyze_wizard(self, ip_column: str = "IP ADDRESS", brand_column: str = "BRAND") -> None:
        """
        Analyze and process the data step-by-step:
        1. Validate IPs.
        2. Filter valid and unique IPs.
        3. Match specific brand patterns.
        4. Extract AREA from VARIABLE.
        5. Format and save result to df_dict["result"].

        :param ip_column: Name of the IP address column.
        :param brand_column: Name of the brand column.
        :raises ValueError: If DataFrame is not loaded.
        """

        self.validate_ipv4_addresses()
        self.split_brand_match()

        if self.df is None:
            raise ValueError("Data not loaded.")

        # Ensure IP column is string and clean whitespace
        self.df[ip_column] = self.df[ip_column].astype(str).str.strip()

        # Step 1: Get valid IP DataFrame
        valid_ip_df = self.df_dict['valid_ips']

        # Step 2: Match brand from valid IPs
        valid_ip_df[brand_column] = valid_ip_df[brand_column].astype(str).str.strip()
        brand_pattern = '|'.join(['SLBS', '2L1T', '3L1T', '4L1T', '2L2T', '3L', '4L'])
        result_df = valid_ip_df[valid_ip_df[brand_column].str.contains(brand_pattern, case=False, na=False)].copy()

        # Step 3: Extract AREA from VARIABLE column
        result_df['AREA'] = result_df['VARIABLE'].astype(str).str.split('_').str[0]

        # Step 4: Insert AREA before TYPE column
        if 'TYPE' in result_df.columns:
            type_index = result_df.columns.get_loc('TYPE')
            cols = result_df.columns.tolist()
            # Move AREA to the correct index
            cols.insert(type_index, cols.pop(cols.index('AREA')))
            result_df = result_df[cols]


        # Step 5: Extract only the matching brand pattern from BRAND column
        # This extracts the first matched pattern string from BRAND

        pattern_regex = re.compile(brand_pattern, re.IGNORECASE)
        
        def extract_brand_pattern(brand_str):
            match = pattern_regex.search(brand_str)
            return match.group(0) if match else brand_str

        result_df[brand_column] = result_df[brand_column].apply(extract_brand_pattern)

        # Step 6: Remove PROJECT and VARIABLE columns before saving
        result_df.drop(columns=["PROJECT", "VARIABLE"], errors="ignore", inplace=True)

        # Step 6: Save result only
        self.df_dict["result"] = result_df

    def export_results(self, output_path: str) -> None:
        """
        Export processed data to an Excel file with multiple sheets.

        :param output_path: Path to the output Excel file.
        :raises ValueError: If there is no data to export.
        """
        if not self.df_dict:
            raise ValueError("No data to export.")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # Step 1: Write 'result' sheet first (if exists)
            if "result" in self.df_dict:
                self.df_dict["result"].to_excel(writer, sheet_name="result", index=False)

            # Step 2: Write all other sheets except 'result' and 'original'
            for key, df in self.df_dict.items():
                if key not in ("result", "original"):
                    df.to_excel(writer, sheet_name=key[:31], index=False)

            # Step 3: Write 'original' sheet last (if exists)
            if "original" in self.df_dict:
                self.df_dict["original"].to_excel(writer, sheet_name="original", index=False)

