import os
import json
import logging


def load_pid_to_skus_map(mapping_path: str | None) -> dict[str, list[str]]:
    """Load PID→SKUs mapping from an external JSON file.

    Expected JSON structure:
    {
      "AIR-DNA-E": ["AIR-DNA-E-T"],
      "DNA-P-T2-E-5Y": ["DSTACK-T2-E"]
    }
    """
    logger = logging.getLogger(__name__)
    if mapping_path is None:
        default_path = os.path.join(os.getcwd(), "sku_map.json")
        if os.path.exists(default_path):
            mapping_path = default_path
        else:
            logger.info("No external mapping provided; proceeding without PID→SKU exceptions.")
            return {}

    try:
        with open(mapping_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("Mapping JSON must be an object mapping strings to list[str]")
        normalized: dict[str, list[str]] = {}
        for key, value in data.items():
            if isinstance(value, list):
                normalized[str(key).strip()] = [str(v).strip() for v in value]
            elif isinstance(value, str):
                normalized[str(key).strip()] = [value.strip()]
            else:
                logger.warning("Ignoring invalid mapping value for key '%s': %r", key, value)
        logger.info("Loaded %d PID→SKU exception entries from %s", len(normalized), mapping_path)
        return normalized
    except FileNotFoundError:
        logger.warning("Mapping file not found: %s. Continuing without exceptions.", mapping_path)
        return {}
    except Exception:
        logger.exception("Failed to load mapping file: %s. Continuing without exceptions.", mapping_path)
        return {}


def get_valid_sku_matches(cssm_matches, pre_ea_migrated_pid_str: str, pid_to_skus_map: dict[str, list[str]]):
    """Return CSSM rows that match the PID directly or via the exception map.

    - First tries a direct SKU match on the PRE-EA migrated PID.
    - If none, looks up mapped exception SKUs from pid_to_skus_map and matches any of them.
    """
    # Try direct match
    sku_match = cssm_matches[cssm_matches['SKU'] == pre_ea_migrated_pid_str]
    # Try mapped exceptions if no direct match
    if sku_match.empty and pre_ea_migrated_pid_str in pid_to_skus_map:
        valid_skus = pid_to_skus_map[pre_ea_migrated_pid_str]
        sku_match = cssm_matches[cssm_matches['SKU'].isin(valid_skus)]
    return sku_match





    # def _load_xml_df(self, path):
    #     """Wrapper to load Excel/XML-like data; kept for compatibility."""
    #     try:
    #         df = self._load_excel_df(path)
    #         return df
    #     except Exception as e:
    #         self.logger.error("Failed to load xml/excel-like file %s: %s", path, e)
    #         raise

    # def _load_excel_df(self, path, sheet_name=None, header=None):
    #     """Generic Excel loader used by other loaders. Trims column names."""
    #     try:
    #         # let pandas handle sheet/header defaults if None
    #         df = pd.read_excel(path, sheet_name=sheet_name, header=header)
    #         # normalize column names (if any)
    #         if df is not None and not df.empty:
    #             df.columns = df.columns.map(lambda c: str(c).strip() if c is not None else c)
    #         self.logger.info("Loaded dataframe from %s (sheet=%s header=%s) rows=%d", path, sheet_name, header, len(df) if df is not None else 0)
    #         return df
    #     except Exception as e:
    #         self.logger.error("Failed to read Excel file %s: %s", path, str(e))
    #         raise

    # def _load_cssm(self, cssm_path):
    #     """Load and normalize the CSSM 'License Detail' sheet."""
    #     try:
    #         cssm = self._load_excel_df(cssm_path, sheet_name='License Detail', header=5)
    #         # ensure expected columns are strings and trimmed
    #         for col in ('Source Identifier', 'SKU'):
    #             if col in cssm.columns:
    #                 cssm[col] = cssm[col].astype(str).str.strip()
    #         self.logger.info("Successfully loaded CSSM from %s", cssm_path)
    #         return cssm
    #     except Exception as e:
    #         self.logger.error("Failed to load CSSM from %s: %s", cssm_path, str(e))
    #         raise

    # def _load_pre_ea(self, pre_ea_path):
    #     """Load and normalize the pre_ea DataFrame."""
    #     try:
    #         pre_ea = self._load_excel_df(pre_ea_path)
    #         self.logger.info("Successfully loaded pre_ea from %s", pre_ea_path)
    #         return pre_ea
    #     except Exception as e:
    #         self.logger.error("Failed to load pre_ea from %s: %s", pre_ea_path, str(e))
    #         raise
    
    # def _prepare_output_workbook(self, pre_ea_path):
    #     """Save a copy of pre_ea to output_dir and return workbook and active worksheet."""
    #     try:
    #         base_name = os.path.basename(pre_ea_path)
    #         out_name = base_name.replace('.xlsx', '_compared.xlsx')
    #         out_path = os.path.join(self.output_dir, out_name)
    #         wb = load_workbook(pre_ea_path)
    #         wb.save(out_path)
    #         wb = load_workbook(out_path)
    #         ws = wb.active
    #         self.logger.info("Successfully prepared output workbook at %s", out_path)
    #         return out_path, wb, ws
    #     except Exception as e:
    #         self.logger.error("Failed to prepare output workbook from %s: %s", pre_ea_path, str(e))
    #         raise

    # def _extract_and_compare_rows(self, pre_ea, cssm):
    #     """
    #     Extract and compare rows from pre_ea and cssm dataframes.
    #     Returns tuples of (red_rows, blue_rows, yellow_rows, green_rows, pink_rows, used_cssm_indices).
    #     """
    #     logger = self.logger
    #     logger.info("Loaded PRE-EA rows: %d | CSSM rows: %d", len(pre_ea), len(cssm))
        
    #     red_rows, blue_rows, yellow_rows, green_rows, pink_rows = 0, 0, 0, 0, 0
    #     used_cssm_indices = set()  # Tracks used CSSM rows to prevent re-matching

    #     for idx, row in pre_ea.iterrows():
    #         alc_order_number = row.get('ALC Order Number')
    #         pre_ea_migrated_pid = row.get('Pre EA Migrated Pid')
    #         pre_ea_qty = row.get('Quantity')
    #         pre_ea_exp = row.get('Expiration Date')
    #         excel_row_idx = idx + 2
            
    #         # Placeholder: comparison logic to be implemented
    #         pass

    #     return red_rows, blue_rows, yellow_rows, green_rows, pink_rows, used_cssm_indices

    # def compare_and_save(self, pre_ea_path, cssm_path):
    #     """
    #     Compare provided pre_ea and cssm files and save a workbook copy of pre_ea to output_dir.
    #     Uses helper methods to keep responsibilities separated.
    #     """
    #     logger = self.logger
    #     logger.info("Starting comparison: pre_ea=%s cssm=%s", pre_ea_path, cssm_path)
    #     start_time = datetime.now()

    #     cssm = self._load_cssm(cssm_path)
    #     out_path, wb, ws = self._prepare_output_workbook(pre_ea_path)
    #     pre_ea = self._load_pre_ea(pre_ea_path)

    #     # Extract and compare rows
    #     red_rows, blue_rows, yellow_rows, green_rows, pink_rows, used_cssm_indices = self._extract_and_compare_rows(pre_ea, cssm)

    #     elapsed = datetime.now() - start_time
    #     logger.info("Comparison finished in %.2f seconds", elapsed.total_seconds())

    #     return out_path, red_rows, blue_rows, yellow_rows, green_rows, pink_rows, elapsed.total_seconds()




    # Run comparison and save
    # Display results
    # print(f"Output file saved at: {out_path}")
    # print(f"Red rows: {red_rows} | Blue rows: {blue_rows} | Yellow rows: {yellow_rows}")
    # print(f"Green rows: {green_rows} | Pink rows: {pink_rows}")
    # print(f"Elapsed time: {elapsed:.2f} seconds")
    # print("="*60 + "\n")
