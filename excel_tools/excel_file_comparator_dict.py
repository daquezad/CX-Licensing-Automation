import os
import sys
import logging
from datetime import datetime
import pandas as pd
import numpy as np
from openpyxl import load_workbook

# Add parent directory to path for relative imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.fs_utils import ensure_clean_dir
from utils.logging_utils import setup_logging
from utils.mapping_utils import load_pid_to_skus_map, get_valid_sku_matches
from utils.date_utils import standardize_date
from utils.colors import RED_FILL, BLUE_FILL, YELLOW_FILL, GREEN_FILL, PINK_FILL
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

class ExcelFileComparator:
    def __init__(self, output_dir=None, log_filename="compare_excels.log"):
        self.output_dir = output_dir or os.path.join(os.getcwd(), "output_files")
        ensure_clean_dir(self.output_dir)
        self.logger = setup_logging(log_dir=self.output_dir, log_filename=log_filename)

    def _load_df(self, file_path, sheet_name=None, header=0):
        """Helper method to load a DataFrame from an Excel file."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
            df.columns = df.columns.map(str.strip)  # Strip whitespace from column names
            return df
        except Exception as e:
            self.logger.error("Error loading file %s: %s", file_path, e)
            raise
    def find_common_items_in_columns(self,
        df1: pd.DataFrame,
        column1_name: str,
        df2: pd.DataFrame,
        column2_name: str
    ) -> list:

        # Input validation: Check if columns exist in their respective DataFrames
        if column1_name not in df1.columns:
            raise ValueError(f"Column '{column1_name}' not found in DataFrame 1.")
        if column2_name not in df2.columns:
            raise ValueError(f"Column '{column2_name}' not found in DataFrame 2.")
        # Extract unique items from each column and convert to sets for efficient comparison
        set1 = set(df1[column1_name].astype(str).str.strip().unique())
        set2 = set(df2[column2_name].astype(str).str.strip().unique())

        # Find the intersection of the two sets
        common_items_set = set1.intersection(set2)
        # Convert the set of common items to a list and return
        return list(common_items_set)

    def compute_licensing_files(self, pre_ea_path, cssm_path, pid_to_skus_map):
        logger = self.logger
        logger.info("Starting comparison: pre_ea=%s cssm=%s", pre_ea_path, cssm_path)
        start_time = datetime.now()

        # Load CSSM data
        logger.info("Loading CSSM data from %s", cssm_path)
        df_cssm = self._load_df(cssm_path, sheet_name='License Detail', header=5)

        # Load PRE-EA data
        logger.info("Loading PRE-EA data from %s", pre_ea_path)
        df_pre_ea = self._load_df(pre_ea_path,sheet_name='PRE_EA_REPORT', header=0)
        df_pre_ea['Flag'] = ''
        df_pre_ea['Logging Info'] = ''

        # Convert the 'Due_Date' column to datetime objects
        df_pre_ea['Expiration Date'] = pd.to_datetime(df_pre_ea['Expiration Date'])
        today = pd.to_datetime(datetime.today().date())
        mask = df_pre_ea['Expiration Date'] < today
        df_pre_ea.loc[mask, 'Flag'] = 'PURPLE'
        df_pre_ea.loc[mask, 'Logging Info'] = "🟪 Expiration date has already expired." 

        common_codes_dictionary = {}
        df_cssm['Used'] = ''

        # Log common matches: df_pre_ea-> 'ALC Order Number' | df_cssm->'Source Identifier'
        try:
            common_codes_id = self.find_common_items_in_columns(df_pre_ea, 'ALC Order Number', df_cssm, 'Source Identifier')
            self.logger.info(f"Common matches: {len(common_codes_id)}")
        except ValueError as e:
            self.logger.error(f"\nError: {e}")
        # FLAG RED for non-matching ALC Order Numbers
        mask = df_pre_ea['ALC Order Number'].astype(str).str.strip().isin(common_codes_id)
        df_pre_ea.loc[~mask, 'Flag'] = 'RED'
        df_pre_ea.loc[~mask, 'Logging Info'] = "🟥 FLAG RED for non-matching ALC Order Numbers"

        mask = df_cssm['Source Identifier'].astype(str).str.strip().isin(common_codes_id)
        df_cssm.loc[~mask, 'Used'] = 'YES'

        # Normalize pre_ea_migrated_pid_str column to valid sku (usig dictionary)
        if pid_to_skus_map is not None:
            df_pre_ea['Pre EA Migrated Pid'] = df_pre_ea['Pre EA Migrated Pid'].map(pid_to_skus_map).fillna(df_pre_ea['Pre EA Migrated Pid'])

        # FLAG RED for non-matching No SKU (or mapped exception)
        try:
            common_skus = self.find_common_items_in_columns(df_pre_ea, 'Pre EA Migrated Pid', df_cssm, 'SKU')
            self.logger.info(f"Common SKU matches: {len(common_skus)}")
            mask = df_pre_ea['Pre EA Migrated Pid'].astype(str).str.strip().isin(common_skus)
            df_pre_ea.loc[~mask, 'Flag'] = 'RED'
            df_pre_ea.loc[~mask, 'Logging Info'] = "🟥 FLAG RED for non-matching SKU (or mapped exception)"
        except ValueError as e:
            self.logger.error(f"\nError: {e}")
        
        dictionary_filter = {}

        for source_identifier in common_codes_id:
            mask_pre_ea = df_pre_ea['ALC Order Number'].astype(str).str.strip() == str(source_identifier).strip()
            mask_flag = df_pre_ea['Flag'] == ''
            combined_mask_1 = mask_pre_ea & mask_flag
            unique_skus = df_pre_ea[combined_mask_1]['Pre EA Migrated Pid'].astype(str).str.strip().unique()

            # Ensure the dictionary entry is a list before appending
            if source_identifier not in dictionary_filter:
                dictionary_filter[source_identifier] = []
            for sku in unique_skus:
                dictionary_filter[source_identifier].append(sku) # Just call append(), no assignment
        
        common_codes_dictionary = {}
        for source_id, sku_list in dictionary_filter.items():
            # print(f"Source Identifier: {source_id}")
            if source_id not in common_codes_dictionary:
                common_codes_dictionary[source_id] = {}
            for sku in sku_list:
                # print(f"    Individual SKU: {sku}")
                mask_sku = df_pre_ea['Pre EA Migrated Pid'].astype(str).str.strip() == str(sku).strip()
                mask_flag = df_pre_ea['Flag'] == ''
                mask_pre_ea = df_pre_ea['ALC Order Number'].astype(str).str.strip() == str(source_id).strip()
                combined_mask = mask_pre_ea & mask_sku & mask_flag

                mask_sku_cssm = df_cssm['SKU'].astype(str).str.strip() == sku
                mask_source_id_cssm = df_cssm['Source Identifier'].astype(str).str.strip() == str(source_id).strip()
                combined_mask_cssm = mask_sku_cssm & mask_source_id_cssm
                common_codes_dictionary[source_id][sku] = (df_pre_ea[combined_mask], df_cssm[combined_mask_cssm])

        # print(common_codes_dictionary["98877762"]["C1A1TN9300XF-5Y"])
# check point

        # Create a loop that will iterate in common_codes_directionary and make some conditions
        ian = 0
        for source_id, skus_and_data in common_codes_dictionary.items():
            for item in common_codes_dictionary[source_id]:
                # print(f"{source_id} saludso -> {skus_and_data[item][1]}")
                pre_ea_row = skus_and_data[item][0]
                cssm = skus_and_data[item][1]
                
                for pre_row in pre_ea_row.iterrows():
                    print(pre_row)
                
            # create condition to check Pre ea migrated pid  and sku in df_cssm_filtered
                # if pre_ea_row['Pre EA Migrated Pid'] not in df_cssm_filtered['SKU'].values:
                #     df_pre_ea.at[idx, 'Flag'] = 'RED'
                #     df_pre_ea.at[idx, 'Logging Info'] = "🟥 FLAG RED: SKU not found in CSSM for this ALC Order Number."
                #     continue    


        # for source_identifier, sku_data in common_codes_dictionary.items():        
        #     for sku, (df_pre_ea_filtered, df_cssm_filtered) in sku_data.items():
        #         if len(df_pre_ea_filtered) != 0 and len(df_cssm_filtered) != 0:
        #             for idx, pre_ea_row in df_pre_ea_filtered.iterrows():
        #             # create condition to check Pre ea migrated pid  and sku in df_cssm_filtered
        #                 if pre_ea_row['Pre EA Migrated Pid'] not in df_cssm_filtered['SKU'].values:
        #                     df_pre_ea.at[idx, 'Flag'] = 'RED'
        #                     df_pre_ea.at[idx, 'Logging Info'] = "🟥 FLAG RED: SKU not found in CSSM for this ALC Order Number."
        #                     continue    
        #             for idx, pre_ea_row in df_pre_ea_filtered.items():
        #                 pre_ea_exp_date_for_check = standardize_date(pre_ea_row['Expiration Date'], in_format="%Y-%m-%d %H:%M:%S")
        #                 if not pre_ea_exp_date_for_check:
        #                     df_pre_ea.at[idx, 'Flag'] = 'YELLOW'
        #                     df_pre_ea.at[idx, 'Logging Info'] = "🟨 FLAG YELLOW: Invalid or empty Expiration Date."
        #                     continue

        #             # Mark rows as used in df_cssm_filtred subscription End date is empty or invalid 
        #                 for cssm_idx, cssm_row in df_cssm_filtered.iterrows():
        #                     cssm_exp = cssm_row['Subscription End Date']
        #                     cssm_exp_date_for_check = standardize_date(cssm_exp, in_format="%Y-%m-%d %H:%M:%S")
        #                     if not cssm_exp_date_for_check:
        #                         df_cssm_filtered.at[cssm_idx, 'Used'] = 'YES'

        #             # clean used and flagged rows
        #             mask_flag = df_cssm_filtered['Used'] == ''
        #             df_cssm_filtered = df_cssm_filtered[mask_flag]
        #             mask_flag = df_pre_ea_filtered['Flag'] == ''
        #             df_pre_ea_filtered = df_pre_ea_filtered[mask_flag]
            
        #             for idx, pre_ea_row in df_pre_ea_filtered.iterrows():
        #                 pre_ea_qty = pre_ea_row['Quantity']
        #                 pre_ea_exp_date_for_check = standardize_date(pre_ea_row['Expiration Date'], in_format="%Y-%m-%d %H:%M:%S")
        #                 for cssm_idx, cssm_row in df_cssm_filtered.iterrows():
        #                     cssm_qty = cssm_row['Available To Use']
        #                     cssm_exp = standardize_date(cssm_row['Subscription End Date'])
        # #                     # check if quantity is equal and date is valid 
        #                     if cssm_qty == pre_ea_qty   and (pre_ea_exp_date_for_check and cssm_exp and pre_ea_exp_date_for_check <= cssm_exp) and cssm_row['Used'] != 'YES':
        #                         df_cssm.at[cssm_idx, 'Used'] = 'YES'
        #                         # flag green
        #                         df_pre_ea.at[idx, 'Flag'] = 'GREEN'
        #                         df_pre_ea.at[idx, 'Logging Info'] = "🟩 FLAG GREEN: Quantity and Expiration Date match found."


                    # # Check leftouvers in the qty coincidences
                    # for idx, pre_ea_row in df_pre_ea_filtered.iterrows():
                    #     # compare quantity of cssm filtred and validate that date is ok
                    #     pre_ea_qty = pre_ea_row['Quantity']
                    #     pre_ea_exp = pre_ea_row['Expiration Date']
                    #     pre_ea_exp_date_for_check = standardize_date(pre_ea_exp, in_format="%Y-%m-%d %H:%M:%S")
                    #     # # Exclude used rows 
                    #     mask_flag = df_cssm['Used'] == ''
                    #     df_cssm_filtered = df_cssm_filtered[mask_flag]
                    
                    #     # get sum of available to use in cssm filtered
                    #     total_cssm_qty = df_cssm_filtered['Available To Use'].sum()
                    #     # get oldest expiration date in cssm filtered
                    #     cssm_filtered_exp_dates = df_cssm_filtered['Subscription End Date'].apply(lambda x: standardize_date(x, in_format="%Y-%m-%d %H:%M:%S"))
                    #     oldest_cssm_exp_date = cssm_filtered_exp_dates.min()

                    #     if total_cssm_qty < pre_ea_qty:
                    #         df_pre_ea.at[idx, 'Flag'] = 'RED'
                    #         df_pre_ea.at[idx, 'Logging Info'] = "🟥 FLAG RED: Cumulative Quantity match NOT found."
                    #     #compare if early cssm before pre_ea expiration date
                    #     elif oldest_cssm_exp_date and pre_ea_exp_date_for_check and oldest_cssm_exp_date < pre_ea_exp_date_for_check:
                    #         df_pre_ea.at[idx, 'Flag'] = 'YELLOW'
                    #         df_pre_ea.at[idx, 'Logging Info'] = "🟨 FLAG YELLOW: Cumulative Quantity match found but Expiration Date issues."
                    #     else:
                    #         df_pre_ea.at[idx, 'Flag'] = 'GREEN'
                    #         df_pre_ea.at[idx, 'Logging Info'] = "🟩 FLAG GREEN: Cumulative Quantity and Expiration Date match found."
                    #         # mark used in cssm filtered
                    #         for cssm_idx, cssm_row in df_cssm_filtered.iterrows():
                    #             df_cssm.at[cssm_idx, 'Used'] = 'YES'
                    #             # counte green flags

    # counte green flags    
        green_count = df_pre_ea['Flag'].value_counts().get('GREEN', 0)
        print(f"Number of GREEN rows: {green_count}")
        red_count = df_pre_ea['Flag'].value_counts().get('RED', 0)
        print(f"Number of RED rows: {red_count}")
        purple_count = df_pre_ea['Flag'].value_counts().get('PURPLE', 0)
        print(f"Number of PURPLE rows: {purple_count}")
        yellow_count = df_pre_ea['Flag'].value_counts().get('YELLOW', 0)
        print(f"Number of YELLOW rows: {yellow_count}")


        print(len(df_pre_ea))
        # Prepare output workbook

        output_dir = self.output_dir
        base_name = os.path.basename(pre_ea_path)
        out_name = base_name.replace('.xlsx', '_compared.xlsx')
        out_path = os.path.join(output_dir, out_name)

        # Save df_pre_ea to the desired output Excel file
        df_pre_ea.to_excel(out_path, index=False)

        # (Optional) If you want to do further processing with openpyxl:
        from openpyxl import load_workbook

        wb = load_workbook(out_path)
        ws = wb.active


        # logger.info("Loaded PRE-EA rows: %d | CSSM rows: %d", len(pre_ea), len(cssm))

        # Initialize counters for row colors
        red_rows, blue_rows, yellow_rows, green_rows, pink_rows = 0, 0, 0, 0, 0

# Test basic functionality
if __name__ == "__main__":
    print("\n" + "="*60)
    print("BASIC FUNCTIONALITY TEST")
    print("="*60)

    # Define file paths
    pre_ea_path = os.path.join(os.getcwd(), "uploaded_pre_ea.xlsx")
    cssm_path = os.path.join(os.getcwd(), "uploaded_cssm.xlsx")
    pid_to_skus_map = load_pid_to_skus_map(os.path.join(os.getcwd(), "sku_map.json"))

    # print(pid_to_skus_map)
    # Initialize comparator
    comparator = ExcelFileComparator()
    comparator.compute_licensing_files(pre_ea_path, cssm_path, pid_to_skus_map)





