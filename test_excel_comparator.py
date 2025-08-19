import sys
import os
import json
from excel_tools.excel_file_comparator import ExcelFileComparator
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def main(pre_ea_path, cssm_path, sku_map_path=None):
    # Load mapping
    if sku_map_path and os.path.exists(sku_map_path):
        with open(sku_map_path, "r", encoding="utf-8") as f:
            pid_to_skus_map = json.load(f)
    else:
        pid_to_skus_map = {}

    # Read Excel files as bytes
    with open(pre_ea_path, "rb") as f:
        pre_ea_bytes = f.read()
    with open(cssm_path, "rb") as f:
        cssm_bytes = f.read()

    comparator = ExcelFileComparator()
    out_path = comparator.compare_and_save(pre_ea_path, cssm_path, pid_to_skus_map)
    print(f"Comparison complete. Output saved to: {out_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python test_excel_comparator.py pre-ea.xlsx cssm.xlsx [sku_map.json]")
        sys.exit(1)
    pre_ea_path = sys.argv[1]
    cssm_path = sys.argv[2]
    sku_map_path = sys.argv[3] if len(sys.argv) > 3 else None
    print(sku_map_path)
    main(pre_ea_path, cssm_path, sku_map_path)
