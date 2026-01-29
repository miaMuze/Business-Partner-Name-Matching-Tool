import pandas as pd
from thefuzz import process, fuzz
import time
import os

# --- File Configuration ---
# Place the Excel file in the same folder as this script
INPUT_FILE = 'data_migration.xlsx'
OUTPUT_FILE = 'Match_Result_Final.xlsx'

def heavy_clean(text):
    """
    Clean and normalize text for better matching accuracy.
    Removes punctuation, spaces, and common business terms.
    """
    if not isinstance(text, str):
        return ""

    # 1. Convert to lowercase, remove periods and spaces
    text = text.replace(".", "").replace(" ", "").lower()

    # 2. Remove common business prefixes/suffixes that vary in spelling
    # Thai terms: Company, Ltd., Limited, Partnership, Public Co., Mr., Mrs., Shop
    # English terms: Co., Ltd., Corp., Inc., LLC, Mr., Mrs., Ms.
    bad_words = [
        # Thai business terms
        "บริษัท", "บจก", "จำกัด", "หจก", "บมจ", "คุณ", "หสน", "นาง", "นาย", "ร้าน",
        # English business terms
        "company", "limited", "ltd", "co", "corp", "corporation",
        "inc", "incorporated", "llc", "plc", "pcl",
        "mr", "mrs", "ms", "miss"
    ]
    for word in bad_words:
        text = text.replace(word, "")

    return text

def run_migration():
    """
    Main function to run the name matching migration process.
    Reads Oracle and SAP data, performs fuzzy matching, and outputs results.
    """
    # Check if input file exists
    if not os.path.exists(INPUT_FILE):
        print(f"[ERROR] File not found: {INPUT_FILE}")
        print(f"        Please place the file in the same folder as this script.")
        return

    print("[INFO] Loading data from Excel...")
    try:
        # Load data from separate sheets
        oracle_df = pd.read_excel(INPUT_FILE, sheet_name='Oracle')
        sap_df = pd.read_excel(INPUT_FILE, sheet_name='SAP')
    except Exception as e:
        print(f"[ERROR] Failed to load data: {e}")
        return

    # Prepare Search Key (combine Name1 + Name2)
    print("[INFO] Cleaning data and preparing search keys...")
    oracle_df['Full_Name'] = oracle_df['Name1'].fillna('') + " " + oracle_df['Name2'].fillna('')
    oracle_df['Search_Key'] = oracle_df['Full_Name'].apply(heavy_clean)

    sap_df['Full_Name'] = sap_df['Name1'].fillna('') + " " + sap_df['Name2'].fillna('')
    sap_df['Search_Key'] = sap_df['Full_Name'].apply(heavy_clean)

    # Store SAP search keys in a list for faster processing
    sap_choices = sap_df['Search_Key'].tolist()

    results = []
    total = len(oracle_df)
    start_time = time.time()

    print(f"[INFO] Starting search for top 5 matches (Total: {total} records)...")

    for i, o_row in oracle_df.iterrows():
        # Show progress every 50 records
        if i % 50 == 0 and i > 0:
            elapsed = time.time() - start_time
            avg_time = elapsed / i
            remaining = avg_time * (total - i)
            print(f"[PROGRESS] {i}/{total} completed | Elapsed: {elapsed/60:.1f} min | Remaining: {remaining/60:.1f} min")

        # Find Top 5 matches using token_sort_ratio
        # This handles word order differences (e.g., "ABC Corp" matches "Corp ABC")
        top_5 = process.extract(o_row['Search_Key'], sap_choices, scorer=fuzz.token_sort_ratio, limit=5)

        res = {
            'Oracle_ID': o_row['ID'],
            'Oracle_Name': o_row['Full_Name']
        }

        # Loop through and store top 5 results
        for j, (match_str, score) in enumerate(top_5):
            # Find the index of match_str in sap_choices
            idx = sap_choices.index(match_str)
            sap_row = sap_df.iloc[idx]
            res[f'Match_{j+1}_BP_Number'] = sap_row['BP_Number']
            res[f'Match_{j+1}_SAP_Name'] = sap_row['Full_Name']
            res[f'Match_{j+1}_Score'] = score

        results.append(res)

    # Save results to Excel
    print("[INFO] Saving results to file...")
    pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)

    end_time = time.time()
    print(f"[SUCCESS] Complete! Total time: {(end_time - start_time)/60:.2f} minutes")
    print(f"[OUTPUT] Results saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_migration()
