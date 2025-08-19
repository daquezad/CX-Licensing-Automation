import os
import json
import streamlit as st
import pandas as pd
from excel_tools.excel_file_comparator import ExcelFileComparator
from utils.mapping_utils import load_pid_to_skus_map

def main():
    import re

    def parse_log_table(log_path):
        log_rows = []
        row_pattern = re.compile(r"Row (\d+): (.*?)(Marking as (🟥 RED|🟦 BLUE|🟨 YELLOW|🟩 GREEN))\.")
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                for line in f:
                    match = row_pattern.search(line)
                    if match:
                        row_num = int(match.group(1))
                        message = match.group(2).strip()
                        color = match.group(4)
                        log_rows.append({
                            "Row": row_num,
                            "Result": color,
                            "Details": message
                        })
        except Exception:
            pass
        return pd.DataFrame(log_rows)
    st.set_page_config(page_title="Excel File Comparator", layout="centered")
    st.title("🧮 Excel File Comparator")
    st.markdown("""
    ## 🚀 How to Use
    1. **Upload your PRE-EA and CSSM Excel files** using the fields below.
    2. **Manage your SKU mapping** in the sidebar. You can upload a mapping file, edit it, and download the updated version.
    3. **Click 'Run comparison'** to process the files. The compared workbook will be available for download if successful.
    4. **Check the logs** for details about the comparison process.
    
    | 🟥 | RED: No match found |
    | 🟦 | BLUE: Quantity mismatch |
    | 🟨 | YELLOW: Date issues |
    | 🟩 | GREEN: All OK |
    """)
    st.caption("Upload PRE-EA and CSSM files, manage SKU mapping, and run comparison.")

    # Sidebar: Tuning & Settings (expander, with label above)
    updated_map = {}
    with st.sidebar.expander("🛠️ SKU Exceptions Handler", expanded=False):
        st.markdown("""
        **SKU Exceptions Handler is used to manage exceptions and alternate names for the same item.**
        If a PRE-EA Migrated PID has different corresponding SKUs in CSSM, you can map them here. This ensures the comparison recognizes equivalent items even if their names differ between files.
        
        Example:
        - PRE-EA Migrated PID: `DNA-P-T2-E-5Y`
        - CSSM SKUs: `DSTACK-T2-E`
        
        Add all valid CSSM SKUs for each PRE-EA PID to ensure accurate matching.
        """)
        mapping_file = st.file_uploader("Upload sku_map.json (optional)", type=["json"], accept_multiple_files=False, key="sku_map_upload")

        # Load or initialize mapping
        if mapping_file is not None:
            try:
                pid_to_skus_map = json.loads(mapping_file.getvalue().decode("utf-8"))
                if not isinstance(pid_to_skus_map, dict):
                    st.error("Invalid JSON structure. Expected an object of PID → [SKUs].")
                    pid_to_skus_map = {}
            except Exception as e:
                st.error(f"Failed to parse JSON: {e}")
                pid_to_skus_map = {}
        else:
            default_json_path = os.path.join(os.getcwd(), "sku_map.json")
            try:
                pid_to_skus_map = load_pid_to_skus_map(default_json_path)
            except Exception:
                pid_to_skus_map = {}

        # Editable mapping table
        st.subheader("📝 Edit mapping")
        editable_rows = []
        for pid, skus in sorted(pid_to_skus_map.items()):
            editable_rows.append({"Pre EA Migrated Pid": pid, "CSSM SKUs (comma-separated)": ", ".join(skus)})
        edited_df = st.data_editor(pd.DataFrame(editable_rows), num_rows="dynamic", use_container_width=True, key="mapping_editor")

        # Convert edited table back to dict
        for _, r in edited_df.iterrows():
            pid = str(r.get("Pre EA Migrated Pid", "")).strip()
            skus_text = str(r.get("CSSM SKUs (comma-separated)", "")).strip()
            if pid == "":
                continue
            sku_list = [s.strip() for s in skus_text.split(",") if s.strip()]
            if sku_list:
                updated_map[pid] = sku_list

        # Save button: only enabled if changes are made
        initial_json_path = os.path.join(os.getcwd(), "sku_map.json")
        # Load initial mapping for comparison
        try:
            with open(initial_json_path, "r") as f:
                initial_map = json.load(f)
        except Exception:
            initial_map = {}

        changes_made = updated_map != initial_map
        if st.button("💾 Save Changes", disabled=not changes_made):
            try:
                with open(initial_json_path, "w") as f:
                    json.dump(updated_map, f, indent=2)
                st.success("SKU mapping saved to sku_map.json.")
            except Exception as e:
                st.error(f"Failed to save mapping: {e}")

        st.download_button(
            label="⬇️ Download updated sku_map.json",
            data=json.dumps(updated_map, indent=2).encode("utf-8"),
            file_name="sku_map.json",
            mime="application/json",
        )
    # If mapping is hidden, load default or empty for comparison
    if not updated_map:
        default_json_path = os.path.join(os.getcwd(), "sku_map.json")
        try:
            updated_map = load_pid_to_skus_map(default_json_path)
        except Exception:
            updated_map = {}

    st.divider()
    st.header("📊 Compare files")
    pre_ea_file = st.file_uploader("Upload PRE-EA Excel (.xlsx)", type=["xlsx"], key="pre_ea")
    cssm_file = st.file_uploader("Upload CSSM Excel (.xlsx)", type=["xlsx"], key="cssm")

    comparison_result = None
    summary_stats = None
    if pre_ea_file and cssm_file:
        if st.button("Run comparison", type="primary"):
            try:
                # Save uploaded files to disk for comparator
                pre_ea_path = os.path.join(os.getcwd(), "uploaded_pre_ea.xlsx")
                cssm_path = os.path.join(os.getcwd(), "uploaded_cssm.xlsx")
                with open(pre_ea_path, "wb") as f:
                    f.write(pre_ea_file.getvalue())
                with open(cssm_path, "wb") as f:
                    f.write(cssm_file.getvalue())
                comparator = ExcelFileComparator()
                out_path, red_rows, blue_rows, yellow_rows, green_rows, elapsed = comparator.compare_and_save(pre_ea_path, cssm_path, updated_map)
                st.success("Comparison complete. Download the result below.")
                with open(out_path, "rb") as f:
                    st.download_button("Download compared workbook", data=f.read(), file_name=os.path.basename(out_path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                comparison_result = out_path
                summary_stats = (red_rows, blue_rows, yellow_rows, green_rows, elapsed)
            except Exception as e:
                st.error(f"Failed to compare files: {e}")

    if summary_stats:
        red_rows, blue_rows, yellow_rows, green_rows, elapsed = summary_stats
        st.divider()
        st.header("📊 Comparison Summary")
        st.markdown(f"""
        **🟥 RED:** {red_rows} &nbsp;&nbsp; **🟦 BLUE:** {blue_rows} &nbsp;&nbsp; **🟨 YELLOW:** {yellow_rows} &nbsp;&nbsp; **🟩 GREEN:** {green_rows}
        
        ⏱️ **Total time:** {elapsed:.2f} seconds
        """)

    output_dir = os.path.join(os.getcwd(), "output_files")
    log_path = os.path.join(output_dir, "compare_excels.log")

    # Show logs as table (row-by-row results) before Run Logs
    df_logs = None
    if os.path.exists(log_path):
        df_logs = parse_log_table(log_path)
        if not df_logs.empty:
            st.divider()
            st.header("🗂️ Row-by-row Results")
            st.dataframe(df_logs, use_container_width=True)

    st.divider()
    st.subheader("Run logs (details and troubleshooting)")
    if st.button("Refresh logs"):
        pass

    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_text = f.read()
        st.text_area("compare_excels.log", value=log_text, height=350)
        st.download_button(
            label="Download compare_excels.log",
            data=log_text,
            file_name="compare_excels.log",
            mime="text/plain",
        )
    except FileNotFoundError:
        st.info("No log file found yet. Run a comparison to generate logs.")
    except Exception as e:
        st.warning(f"Could not read log file: {e}")

if __name__ == "__main__":
    main()
