"""
Script to create sample English data for data_migration.xlsx
Run this script to generate example data in English

Usage:
    python create_sample_data.py
"""

import pandas as pd

# Sample Oracle data (English)
oracle_data = {
    'ID': [
        'ORC001', 'ORC002', 'ORC003', 'ORC004', 'ORC005',
        'ORC006', 'ORC007', 'ORC008', 'ORC009', 'ORC010',
        'ORC011', 'ORC012', 'ORC013', 'ORC014', 'ORC015'
    ],
    'Name1': [
        'Global Tech', 'Bangkok Electronics', 'Sunrise Trading',
        'Pacific Ocean', 'Diamond Star', 'Golden Dragon',
        'Silver Moon', 'Blue Sky', 'Green Valley', 'Red Mountain',
        'Crystal Clear', 'Northern Light', 'Southern Cross',
        'Eastern Wind', 'Western Sun'
    ],
    'Name2': [
        'Solutions Co., Ltd.', 'Corporation', 'Company Limited',
        'Industries Ltd.', 'Enterprise', 'Holdings Co., Ltd.',
        'International', 'Services Ltd.', 'Group', 'Partners',
        'Technologies', 'Logistics Co., Ltd.', 'Manufacturing',
        'Consulting', 'Distribution Ltd.'
    ]
}

# Sample SAP data (English - with slight variations to test fuzzy matching)
sap_data = {
    'BP_Number': [
        'SAP10001', 'SAP10002', 'SAP10003', 'SAP10004', 'SAP10005',
        'SAP10006', 'SAP10007', 'SAP10008', 'SAP10009', 'SAP10010',
        'SAP10011', 'SAP10012', 'SAP10013', 'SAP10014', 'SAP10015',
        'SAP10016', 'SAP10017', 'SAP10018', 'SAP10019', 'SAP10020'
    ],
    'Name1': [
        # Exact or near matches
        'Global Tech', 'Bangkok Electronic', 'Sunrise Trade',
        'Pacific Ocean', 'Dimond Star', 'Golden Dragn',
        'Silver Moon', 'Blue Sky', 'Green Valey', 'Red Mountan',
        # Different companies
        'Alpha Beta', 'Omega Systems', 'Delta Force', 'Gamma Ray',
        'Theta Wave', 'Crystal Clar', 'Northen Light', 'Southern Cros',
        'Eastern Wynd', 'Western Sunn'
    ],
    'Name2': [
        # Matching variations
        'Solutions Ltd.', 'Corp.', 'Co. Ltd.',
        'Industries', 'Enterprise Co., Ltd.', 'Holdings Ltd.',
        'Intl.', 'Services', 'Group Co., Ltd.', 'Partners Ltd.',
        # Different companies
        'Technologies', 'Solutions', 'Security', 'Energy',
        'Communications', 'Tech', 'Logistics Ltd.', 'Mfg.',
        'Consultants', 'Dist.'
    ]
}

# Create DataFrames
oracle_df = pd.DataFrame(oracle_data)
sap_df = pd.DataFrame(sap_data)

# Save to Excel with two sheets
output_file = 'data_migration.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    oracle_df.to_excel(writer, sheet_name='Oracle', index=False)
    sap_df.to_excel(writer, sheet_name='SAP', index=False)

print(f"Sample data created successfully!")
print(f"Output file: {output_file}")
print(f"\nOracle records: {len(oracle_df)}")
print(f"SAP records: {len(sap_df)}")
print(f"\n--- Oracle Data Preview ---")
print(oracle_df.to_string(index=False))
print(f"\n--- SAP Data Preview ---")
print(sap_df.to_string(index=False))
