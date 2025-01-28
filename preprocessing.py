import pandas as pd
import numpy as np

# Load the original data from Sheet2 of the Excel file (replace filename)
df = pd.read_excel('test.xlsx', sheet_name='Sheet1')

def consolidate_data(df):
    # 1. Group by MARK and PCS/CTN, then count cartons and sum quantities.
    consolidated = df.groupby(['MARK', 'PCS/CTN', 'DESCRIPTION'])['CTN NO'].count().reset_index()
    consolidated.rename(columns={'CTN NO': 'T.CTN', 'PCS/CTN': 'QTY'}, inplace=True)

    consolidated['T.QTY'] = consolidated['T.CTN'] * consolidated['QTY']

    # 2. Handle Units
    def handle_units(row):
        if row['MARK'] == 'VRB50' and 'WATCH CELL' in row['DESCRIPTION']:
            return 'PCS'
        elif row['MARK'] == 'VRB50' and 'METAL BUTTON' in row['DESCRIPTION']:
            return 'KGS'  # Updating Unit for Metal Buttons
        return 'PCS'

    consolidated['UNIT'] = consolidated.apply(handle_units, axis=1)

    # # 3. Handle Weight (WT) – Careful averaging/first-value logic
    # def get_weight(group):
    #     weights = df[(df['MARK'] == group.name[0]) & (df['PCS/CTN'] == group.name[1]) & (df['DESCRIPTION'] == group.name[2])]['WEIGHT/TOTAL'].dropna().unique()
    #     if len(weights) == 1:
    #         return weights[0] #return total weight if possible
    #     elif len(weights) > 1:
    #         return None  # Indicate inconsistent weights (handle later)
    #     return None  # Missing weight


    consolidated['WT'] = consolidated.apply(get_weight, axis=1)


    # 4. Consolidate CTN NO ranges
    def consolidate_ctn_no(group):
        ctn_nos = sorted(df[(df['MARK'] == group.name[0]) & (df['PCS/CTN'] == group.name[1]) & (df['DESCRIPTION'] == group.name[2])]['CTN NO'].dropna().unique())
        if len(ctn_nos) == 1:
            return str(ctn_nos[0])
        else:
            start = ctn_nos[0]
            end = ctn_nos[-1]

            if all(isinstance(n, int) or (isinstance(n, str) and n.isdigit())for n in ctn_nos):
                    return f"{start}-{end}"  # Numeric range
            elif len(ctn_nos) < 4:
                return ', '.join(map(str, ctn_nos))
            else:
                return f"{start}-{end}"  # default string range



    consolidated['CTN NO'] = consolidated.apply(consolidate_ctn_no, axis=1)

    return consolidated

def get_weight(group):
    # Access the MARK, PCS/CTN, and DESCRIPTION directly from the row (series)
    mark = group['MARK']
    pcs_ctn = group['QTY']  # Use 'QTY' which is the new name for PCS/CTN
    description = group['DESCRIPTION']

    weights = df[(df['MARK'] == mark) & (df['PCS/CTN'] == pcs_ctn) & (df['DESCRIPTION'] == description)]['WEIGHT/TOTAL'].dropna().unique()

    if len(weights) == 1:
        return weights[0]
    elif len(weights) > 1:
        return np.mean(weights)  # Or handle differently if needed
    return None  # Explicitly return None for missing weights

# ... (rest of your consolidate_data function and other code)

consolidated_df = consolidate_data(df)

#Post Consolidation weight adjustments
kgs_rows = consolidated_df[consolidated_df['UNIT'] == 'KGS']
consolidated_df.loc[kgs_rows.index, 'WT'] = kgs_rows['QTY'] #Ensuring weight matches qty when kgs is unit

consolidated_df['WT'] = consolidated_df.apply(lambda row: row['WT'] / row['T.CTN'] if row['UNIT'] == 'PCS' else row['WT'], axis=1).fillna('')  #Weight adjustments based on units

# 5. Standardize Descriptions (refine as needed!)
def standardize_description(desc):
    desc = desc.upper()
    if "BACK COVER" in desc:
        return "PLASTIC BACK COVER (FOR MOBILE)"
    elif "BATTERY" in desc:
        return "MOBILE BATTERY (R-41206857)"
    elif "PLOYBAG" in desc or "PACKING  BOX" in desc:
        return "PACKING MATERIAL"
    elif "LCD FRAME 中框" in desc:
        return "MIDDLE FRAME (FOR MOBILE HOUSING)"
    elif "LCD" in desc:
        return "LCD DISPLAY (FOR MOBILE)"
    elif "UNLOCK MAGNET" in desc:
        return "KEYCHAIN"
    elif "PHONE HOLDER" in desc:
        return "TRIPOD STAND"
    elif "PHONE STAND" in desc:
        return "UNIVARSAL TAB HOLDER"
    elif "MAT" in desc:
        return "MINI MAT (FOR MOBILE)"
    elif "WATCH CELL" in desc:
        return "METAL BUTTON"
    return desc

consolidated_df['DESCRIPTION'] = consolidated_df['DESCRIPTION'].apply(standardize_description)

#... (Other cleaning/formatting) ...

# Save to CSV
print(consolidated_df)
consolidated_df.to_csv('consolidated_data.csv', index=False)