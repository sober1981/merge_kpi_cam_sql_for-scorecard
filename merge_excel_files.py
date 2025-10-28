"""
Excel Files Merger Script
This script merges multiple Excel files into a single file with standardized headers
based on the FORMAT GRAL TABLE mapping.

Author: Created for drilling optimization project
Date: 2025-10-28
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURATION
# ============================================================================

# File paths
FILES = {
    'Motor_KPI': 'Motor KPI (16).xlsx',
    'CAM_Run_Tracker': 'CAM Run Tracker Rev 4 (14)_example.xlsx',
    'POG_CAM_Usage': 'POG CAM Usage (2).xlsx',
    'POG_MM_Usage': 'POG MM Usage (3).xlsx'
}

MAPPING_FILE = 'FORMAT GRAL TABLE.xlsx'
BASIN_LOOKUP_FILE = 'LISTS_BASIN AND FORM_FAM.xlsx'
OUTPUT_FILE = f'MERGED_DATA_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

# Operator name standardization mapping (CAM Run Tracker -> Standard names)
OPERATOR_MAPPING = {
    'Aethon Energy': 'Aethon Energy Operating, LLC',
    'BPX': 'BPX Operating Company',
    'COMSTOCK RESOURCES': 'Comstock Oil & Gas LLP',
    'Camino': 'Camino Resources',
    'Caturus Energy': 'CATURUS ENERGY, LLC',
    'Comstock': 'Comstock Oil & Gas LLP',
    'Comstock Resources': 'Comstock Oil & Gas LLP',
    'Conoco': 'Conoco Phillips',
    'ConocoPhillips': 'Conoco Phillips',
    'Coterra': 'COTERRA',
    'Devon': 'Devon Energy',
    'Discovery': 'DISCOVERY NATURAL RESOURCES',
    'Exxon': 'EXXON',
    'Fervo': 'FERVO ENERGY COMPANY',
    'Greenlake Energy': 'GREENLAKE ENERGY',
    'Logos Operating LLC': 'LOGOS OPERATING LLC',
    'Mewbourne': 'Mewbourne Oil Company',
    'Mitsui': 'MITSUI E&P USA LLC',
    'Ovintiv': 'Ovintiv USA',
    'Oxy': 'OXY USA',
    'Oxy EOR': 'OXY USA',
    'Petro-Hunt': 'PETRO-HUNT',
    'Summit': 'Summit Petroleum',
    'XTO': 'EXXON',
}

# ============================================================================
# STEP 1: Load Mapping Configuration
# ============================================================================

def load_mapping():
    """Load the header mapping from FORMAT GRAL TABLE"""
    print("Loading mapping configuration...")

    df_mapping = pd.read_excel(MAPPING_FILE, sheet_name='Sheet1')

    # Create mapping dictionary for each source
    mappings = {}
    target_headers = list(df_mapping.columns)

    for idx, row in df_mapping.iterrows():
        source_name = row['SOURCE']
        mapping = {}

        for target_col in target_headers:
            if target_col == 'SOURCE':
                continue
            source_col = row[target_col]

            # Only map if not "Not Present"
            if pd.notna(source_col) and 'Not Present' not in str(source_col):
                mapping[source_col] = target_col

        mappings[source_name] = mapping

    print(f"  Loaded mappings for {len(mappings)} source files")
    for source, mapping in mappings.items():
        print(f"    {source}: {len(mapping)} mapped columns")

    return mappings, target_headers

# ============================================================================
# STEP 2: Load Lookup Tables
# ============================================================================

def load_lookup_tables():
    """Load Basin and Formation Family lookup tables"""
    print("\nLoading lookup tables...")

    # Load Basin lookup
    basin_df = pd.read_excel(BASIN_LOOKUP_FILE, sheet_name='Basin')

    # Create a dictionary mapping county to basin
    county_to_basin = {}
    for col in basin_df.columns:
        basin_name = col
        for county in basin_df[col].dropna():
            county_to_basin[str(county).strip().upper()] = basin_name

    print(f"  Loaded {len(county_to_basin)} county-to-basin mappings")

    # Load Formation Family lookup
    formfam_df = pd.read_excel(BASIN_LOOKUP_FILE, sheet_name='FORM_FAM')
    print(f"  Loaded {len(formfam_df)} formation family mappings")

    return county_to_basin, formfam_df

# ============================================================================
# STEP 3: Read and Transform Source Files
# ============================================================================

def read_motor_kpi(file_path, mapping):
    """Read Motor KPI file"""
    print(f"\nReading Motor KPI file...")
    df = pd.read_excel(file_path, sheet_name='Motor KPI')
    print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")

    # Save original BHA column before renaming (if it exists)
    original_bha = None
    if 'BHA' in df.columns:
        original_bha = df['BHA'].copy()

    # Rename columns according to mapping
    df_renamed = df.rename(columns=mapping)
    df_renamed['SOURCE'] = 'Motor_KPI'

    # Restore BHA column if it was present and got mapped to something else
    if original_bha is not None:
        if 'BHA' not in df_renamed.columns or df_renamed['BHA'].isna().all():
            df_renamed['BHA'] = original_bha

    # Special handling: Ensure DATEIN and DATEOUT map to DATE_IN and DATE_OUT
    if 'DATEIN' in df.columns and ('DATE_IN' not in df_renamed.columns or df_renamed['DATE_IN'].isna().all()):
        df_renamed['DATE_IN'] = df['DATEIN']
    if 'DATEOUT' in df.columns and ('DATE_OUT' not in df_renamed.columns or df_renamed['DATE_OUT'].isna().all()):
        df_renamed['DATE_OUT'] = df['DATEOUT']

    # Special handling: Map BENDANGLE to BEND column
    if 'BENDANGLE' in df.columns and ('BEND' not in df_renamed.columns or df_renamed['BEND'].isna().all()):
        df_renamed['BEND'] = df['BENDANGLE']
        # Also populate BEND_HSG with the same value if it's empty
        if 'BEND_HSG' not in df_renamed.columns or df_renamed['BEND_HSG'].isna().all():
            df_renamed['BEND_HSG'] = df['BENDANGLE']

    return df_renamed

def read_cam_run_tracker(file_path, mapping):
    """Read CAM Run Tracker file"""
    print(f"\nReading CAM Run Tracker file...")
    df = pd.read_excel(file_path, sheet_name='General')
    print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")

    # Rename columns according to mapping
    df_renamed = df.rename(columns=mapping)
    df_renamed['SOURCE'] = 'CAM_Run_Tracker'

    # Special handling: Map 'Run #' to BHA column (if not already mapped)
    if 'Run #' in df.columns and 'BHA' not in df_renamed.columns:
        df_renamed['BHA'] = df['Run #']
    elif 'Run #' in df.columns and df_renamed['BHA'].isna().all():
        df_renamed['BHA'] = df['Run #']

    # Special handling: Split 'Start of Run' into DATE_IN and TIME_IN
    if 'Start of Run' in df.columns:
        # Extract date part
        if 'DATE_IN' not in df_renamed.columns or df_renamed['DATE_IN'].isna().all():
            df_renamed['DATE_IN'] = pd.to_datetime(df['Start of Run'], errors='coerce').dt.date
        # Extract time part
        if 'TIME_IN' not in df_renamed.columns or df_renamed['TIME_IN'].isna().all():
            df_renamed['TIME_IN'] = pd.to_datetime(df['Start of Run'], errors='coerce').dt.time

    # Special handling: Split 'End of Run' into DATE_OUT and TIME_OUT
    if 'End of Run' in df.columns:
        # Extract date part
        if 'DATE_OUT' not in df_renamed.columns or df_renamed['DATE_OUT'].isna().all():
            df_renamed['DATE_OUT'] = pd.to_datetime(df['End of Run'], errors='coerce').dt.date
        # Extract time part
        if 'TIME_OUT' not in df_renamed.columns or df_renamed['TIME_OUT'].isna().all():
            df_renamed['TIME_OUT'] = pd.to_datetime(df['End of Run'], errors='coerce').dt.time

    # Special handling: Map 'Bend' to BEND column
    if 'Bend' in df.columns and ('BEND' not in df_renamed.columns or df_renamed['BEND'].isna().all()):
        df_renamed['BEND'] = df['Bend']
        # Also populate BEND_HSG with the same value if it's empty
        if 'BEND_HSG' not in df_renamed.columns or df_renamed['BEND_HSG'].isna().all():
            df_renamed['BEND_HSG'] = df['Bend']

    return df_renamed

def read_pog_cam_usage(file_path, mapping):
    """Read POG CAM Usage file"""
    print(f"\nReading POG CAM Usage file...")
    df = pd.read_excel(file_path, sheet_name='POG Tool Usage')
    print(f"  Rows before cleaning: {len(df)}")

    # The first row contains the actual headers
    new_headers = df.iloc[0]
    df = df[1:].copy()
    df.columns = new_headers

    # Remove any completely empty rows
    df = df.dropna(how='all')

    print(f"  Rows after cleaning: {len(df)}, Columns: {len(df.columns)}")

    # Rename columns according to mapping
    df_renamed = df.rename(columns=mapping)
    df_renamed['SOURCE'] = 'POG_CAM_Usage'

    # Special handling: Map Brt Date and Art Date to DATE_IN and DATE_OUT
    if 'Brt Date' in df.columns:
        if 'DATE_IN' not in df_renamed.columns or df_renamed['DATE_IN'].isna().all():
            df_renamed['DATE_IN'] = pd.to_datetime(df['Brt Date'], errors='coerce').dt.date
    if 'Art Date' in df.columns:
        if 'DATE_OUT' not in df_renamed.columns or df_renamed['DATE_OUT'].isna().all():
            df_renamed['DATE_OUT'] = pd.to_datetime(df['Art Date'], errors='coerce').dt.date

    # Special handling: Map Fixed or Adjustable to BEND (use whichever has a value)
    if 'Fixed' in df.columns or 'Adjustable' in df.columns:
        if 'BEND' not in df_renamed.columns or df_renamed['BEND'].isna().all():
            # Try Fixed first, then Adjustable as fallback
            if 'Fixed' in df.columns:
                df_renamed['BEND'] = df['Fixed']
            if 'Adjustable' in df.columns:
                # Fill BEND with Adjustable where Fixed is empty
                df_renamed['BEND'] = df_renamed['BEND'].fillna(df['Adjustable'])

        # Also populate BEND_HSG with the same values
        if 'BEND_HSG' not in df_renamed.columns or df_renamed['BEND_HSG'].isna().all():
            df_renamed['BEND_HSG'] = df_renamed['BEND']

    # Special handling: Map Job Type column to JOB_TYPE
    if 'Job Type' in df.columns:
        if 'JOB_TYPE' not in df_renamed.columns or df_renamed['JOB_TYPE'].isna().all():
            df_renamed['JOB_TYPE'] = df['Job Type']

    return df_renamed

def read_pog_mm_usage(file_path, mapping):
    """Read POG MM Usage file"""
    print(f"\nReading POG MM Usage file...")
    df = pd.read_excel(file_path, sheet_name='POG Tool Usage')
    print(f"  Rows before cleaning: {len(df)}")

    # The first row contains the actual headers
    new_headers = df.iloc[0]
    df = df[1:].copy()
    df.columns = new_headers

    # Remove any completely empty rows
    df = df.dropna(how='all')

    print(f"  Rows after cleaning: {len(df)}, Columns: {len(df.columns)}")

    # Rename columns according to mapping
    df_renamed = df.rename(columns=mapping)
    df_renamed['SOURCE'] = 'POG_MM_Usage'

    # Special handling: Map Brt Date and Art Date to DATE_IN and DATE_OUT
    if 'Brt Date' in df.columns:
        if 'DATE_IN' not in df_renamed.columns or df_renamed['DATE_IN'].isna().all():
            df_renamed['DATE_IN'] = pd.to_datetime(df['Brt Date'], errors='coerce').dt.date
    if 'Art Date' in df.columns:
        if 'DATE_OUT' not in df_renamed.columns or df_renamed['DATE_OUT'].isna().all():
            df_renamed['DATE_OUT'] = pd.to_datetime(df['Art Date'], errors='coerce').dt.date

    # Special handling: Map Fixed or Adjustable to BEND (use whichever has a value)
    if 'Fixed' in df.columns or 'Adjustable' in df.columns:
        if 'BEND' not in df_renamed.columns or df_renamed['BEND'].isna().all():
            # Try Fixed first, then Adjustable as fallback
            if 'Fixed' in df.columns:
                df_renamed['BEND'] = df['Fixed']
            if 'Adjustable' in df.columns:
                # Fill BEND with Adjustable where Fixed is empty
                df_renamed['BEND'] = df_renamed['BEND'].fillna(df['Adjustable'])

        # Also populate BEND_HSG with the same values
        if 'BEND_HSG' not in df_renamed.columns or df_renamed['BEND_HSG'].isna().all():
            df_renamed['BEND_HSG'] = df_renamed['BEND']

    # Special handling: Map Job Type column to JOB_TYPE
    if 'Job Type' in df.columns:
        if 'JOB_TYPE' not in df_renamed.columns or df_renamed['JOB_TYPE'].isna().all():
            df_renamed['JOB_TYPE'] = df['Job Type']

    return df_renamed

# ============================================================================
# STEP 4: Clean County Names
# ============================================================================

def clean_county_names(df, source_name):
    """
    Clean county names by removing 'County', 'Parish', and state abbreviations
    For Motor KPI and POG files: Extract state from county name FIRST, then clean
    Applies to Motor KPI, POG CAM Usage, and POG MM Usage
    """
    if 'COUNTY' not in df.columns:
        return df

    # Only apply to Motor KPI and POG files
    if source_name in ['Motor_KPI', 'POG_CAM_Usage', 'POG_MM_Usage']:
        print(f"  Extracting STATE from county and cleaning county names...")

        import re

        def extract_state_and_clean(county_str):
            """Extract state abbreviation and return (state, clean_county)"""
            if pd.isna(county_str):
                return (None, county_str)

            county_str = str(county_str).strip()
            state = None

            # Extract state abbreviations (2 capital letters at the end)
            # Pattern: " TX", " LA", " OK", " NM", " WY", etc.
            match = re.search(r'\s+([A-Z]{2})$', county_str)
            if match:
                state = match.group(1)
                # Remove the state from county string
                county_str = re.sub(r'\s+[A-Z]{2}$', '', county_str)

            # Remove "County" word (case insensitive)
            county_str = re.sub(r'\s+County\s*', ' ', county_str, flags=re.IGNORECASE)

            # Remove "Parish" word (for Louisiana)
            county_str = re.sub(r'\s+Parish\s*', ' ', county_str, flags=re.IGNORECASE)

            # Clean up extra spaces
            county_str = ' '.join(county_str.split())

            return (state, county_str.strip())

        # Apply extraction and cleaning
        results = df['COUNTY'].apply(extract_state_and_clean)

        # Separate states and cleaned counties
        states = [r[0] for r in results]
        cleaned_counties = [r[1] for r in results]

        # Update STATE column only if it doesn't already have a value
        if 'STATE' not in df.columns:
            df['STATE'] = states
        else:
            # Only fill STATE if currently empty
            df['STATE'] = df['STATE'].fillna(pd.Series(states))

        # Update COUNTY with cleaned names
        df['COUNTY'] = cleaned_counties

        states_extracted = sum(1 for s in states if s is not None)
        print(f"    Extracted STATE for {states_extracted} records")
        print(f"    Cleaned {len(cleaned_counties)} county names")

        # Show sample
        sample_df = df[['COUNTY', 'STATE']].dropna(subset=['COUNTY']).head(5)
        if len(sample_df) > 0:
            print(f"    Sample results:")
            for idx, row in sample_df.iterrows():
                print(f"      County: {row['COUNTY']}, State: {row['STATE']}")

    return df

# ============================================================================
# STEP 5: Standardize Operator Names
# ============================================================================

def standardize_operator_names(df, source_name):
    """Standardize operator names for CAM Run Tracker data"""
    if 'OPERATOR' not in df.columns:
        return df

    # Only apply mapping to CAM Run Tracker data
    if source_name == 'CAM_Run_Tracker':
        print(f"  Standardizing operator names...")

        # Count changes before
        original_names = df['OPERATOR'].value_counts()
        changes_made = 0

        # Apply mapping
        for old_name, new_name in OPERATOR_MAPPING.items():
            mask = df['OPERATOR'] == old_name
            count = mask.sum()
            if count > 0:
                df.loc[mask, 'OPERATOR'] = new_name
                changes_made += 1
                print(f"    {old_name} -> {new_name} ({count} records)")

        print(f"  Total operator names standardized: {changes_made}")

    return df

# ============================================================================
# STEP 6: Format Dates and Create START_DATE/END_DATE
# ============================================================================

def format_dates_and_datetimes(df):
    """
    Format DATE_IN and DATE_OUT to proper date format
    Create START_DATE and END_DATE by combining date + time
    """

    # Convert DATE_IN and DATE_OUT to proper datetime.date format
    if 'DATE_IN' in df.columns:
        df['DATE_IN'] = pd.to_datetime(df['DATE_IN'], errors='coerce').dt.date
    if 'DATE_OUT' in df.columns:
        df['DATE_OUT'] = pd.to_datetime(df['DATE_OUT'], errors='coerce').dt.date

    # Create START_DATE by combining DATE_IN + TIME_IN
    if 'DATE_IN' in df.columns and 'TIME_IN' in df.columns:
        def combine_datetime(row):
            if pd.notna(row['DATE_IN']) and pd.notna(row['TIME_IN']):
                try:
                    # Convert date to datetime
                    date_part = pd.to_datetime(row['DATE_IN'])
                    # Convert time to datetime
                    time_part = pd.to_datetime(str(row['TIME_IN']), format='%H:%M:%S', errors='coerce')

                    if pd.notna(date_part) and pd.notna(time_part):
                        # Combine date and time
                        combined = pd.Timestamp(
                            year=date_part.year,
                            month=date_part.month,
                            day=date_part.day,
                            hour=time_part.hour,
                            minute=time_part.minute,
                            second=time_part.second
                        )
                        return combined
                except:
                    pass
            return None

        # For Motor KPI, create START_DATE from DATE_IN + TIME_IN
        motor_mask = df['SOURCE'] == 'Motor_KPI'
        df.loc[motor_mask, 'START_DATE'] = df[motor_mask].apply(combine_datetime, axis=1)

        # For POG files, use DATE_IN as START_DATE (they don't have TIME_IN)
        pog_mask = df['SOURCE'].isin(['POG_CAM_Usage', 'POG_MM_Usage'])
        df.loc[pog_mask, 'START_DATE'] = pd.to_datetime(df.loc[pog_mask, 'DATE_IN'], errors='coerce')

    # Create END_DATE by combining DATE_OUT + TIME_OUT
    if 'DATE_OUT' in df.columns and 'TIME_OUT' in df.columns:
        def combine_datetime_out(row):
            if pd.notna(row['DATE_OUT']) and pd.notna(row['TIME_OUT']):
                try:
                    # Convert date to datetime
                    date_part = pd.to_datetime(row['DATE_OUT'])
                    # Convert time to datetime
                    time_part = pd.to_datetime(str(row['TIME_OUT']), format='%H:%M:%S', errors='coerce')

                    if pd.notna(date_part) and pd.notna(time_part):
                        # Combine date and time
                        combined = pd.Timestamp(
                            year=date_part.year,
                            month=date_part.month,
                            day=date_part.day,
                            hour=time_part.hour,
                            minute=time_part.minute,
                            second=time_part.second
                        )
                        return combined
                except:
                    pass
            return None

        # For Motor KPI, create END_DATE from DATE_OUT + TIME_OUT
        motor_mask = df['SOURCE'] == 'Motor_KPI'
        df.loc[motor_mask, 'END_DATE'] = df[motor_mask].apply(combine_datetime_out, axis=1)

        # For POG files, use DATE_OUT as END_DATE (they don't have TIME_OUT)
        pog_mask = df['SOURCE'].isin(['POG_CAM_Usage', 'POG_MM_Usage'])
        df.loc[pog_mask, 'END_DATE'] = pd.to_datetime(df.loc[pog_mask, 'DATE_OUT'], errors='coerce')

    print(f"  Formatted DATE_IN/DATE_OUT to date format")
    print(f"  Created START_DATE/END_DATE from date+time combinations")

    return df

# ============================================================================
# STEP 7: Populate LOBE/STAGE and DDS Columns
# ============================================================================

def populate_lobe_stage_and_dds(df):
    """
    Populate LOBE/STAGE column by combining LOBES and STAGES
    Populate DDS column based on source-specific logic
    """

    # Handle LOBE/STAGE column
    if 'LOBE/STAGE' in df.columns and 'LOBES' in df.columns and 'STAGES' in df.columns:
        def combine_lobe_stage(row):
            # For Motor KPI and POG files, combine LOBES + ":" + STAGES
            if row['SOURCE'] in ['Motor_KPI', 'POG_CAM_Usage', 'POG_MM_Usage']:
                lobe = row['LOBES']
                stage = row['STAGES']
                if pd.notna(lobe) and pd.notna(stage):
                    return f"{lobe}:{stage}"
            # CAM Run Tracker: replace "-" with ":" to match format
            elif row['SOURCE'] == 'CAM_Run_Tracker':
                if pd.notna(row['LOBE/STAGE']):
                    return str(row['LOBE/STAGE']).replace('-', ':')
            return row['LOBE/STAGE']

        df['LOBE/STAGE'] = df.apply(combine_lobe_stage, axis=1)
        print(f"  Combined LOBES and STAGES into LOBE/STAGE column")

    # Handle DDS column
    if 'DDS' in df.columns:
        def populate_dds(row):
            source = row['SOURCE']

            # Motor KPI: Always "SDT"
            if source == 'Motor_KPI':
                return 'SDT'

            # CAM Run Tracker: Extract first complete word (company name)
            elif source == 'CAM_Run_Tracker':
                if pd.notna(row['DDS']):
                    dds_value = str(row['DDS']).strip()
                    # Extract first word before space or /
                    import re
                    match = re.match(r'^([A-Za-z]+)', dds_value)
                    if match:
                        return match.group(1)
                return row['DDS']

            # POG files: Based on JOB_TYPE column
            elif source in ['POG_CAM_Usage', 'POG_MM_Usage']:
                if 'JOB_TYPE' in row.index and pd.notna(row['JOB_TYPE']):
                    job_type = str(row['JOB_TYPE']).strip().upper()
                    if 'DIRECTIONAL' in job_type:
                        return 'SDT'
                    elif 'RENTAL' in job_type:
                        return 'Other'
                return None

            return row['DDS'] if pd.notna(row['DDS']) else None

        df['DDS'] = df.apply(populate_dds, axis=1)
        print(f"  Populated DDS column based on source-specific logic")

    return df

# ============================================================================
# STEP 10: Populate TOTAL HRS
# ============================================================================

def populate_total_hrs(df):
    """
    Populate Total Hrs (C+D) column:
    - Motor KPI: CIRC_HOURS + DRILLING_HOURS
    - CAM Run Tracker: Already populated
    - POG files: Already populated
    """
    total_hrs_col = 'Total Hrs (C+D)'
    if total_hrs_col in df.columns:
        def calculate_total_hrs(row):
            # For Motor KPI, sum CIRC_HOURS and DRILLING_HOURS
            if row['SOURCE'] == 'Motor_KPI':
                circ = row.get('CIRC_HOURS', 0) if pd.notna(row.get('CIRC_HOURS')) else 0
                drilling = row.get('DRILLING_HOURS', 0) if pd.notna(row.get('DRILLING_HOURS')) else 0
                return circ + drilling
            # For other sources, keep existing value
            return row[total_hrs_col]

        df[total_hrs_col] = df.apply(calculate_total_hrs, axis=1)
        print(f"  Calculated Total Hrs (C+D) from CIRC_HOURS + DRILLING_HOURS for Motor KPI")

    return df

# ============================================================================
# STEP 11: Add UPDATE Column
# ============================================================================

def add_update_column(df):
    """
    Add UPDATE column with today's date (date when merge is performed)
    """
    from datetime import datetime

    if 'UPDATE' in df.columns:
        df['UPDATE'] = datetime.now().date()
        print(f"  Added UPDATE column with merge date: {datetime.now().date()}")

    return df

# ============================================================================
# STEP 12: Populate MOTOR_TYPE2
# ============================================================================

def populate_motor_type2(df):
    """
    Populate MOTOR_TYPE2 column based on source-specific logic:

    Motor KPI:
    - If "MLA07" in SN -> "CAM DD"
    - If "TDI" in MOTOR MAKE and no "MLA07" in SN -> "TDI CONV"
    - If no "TDI" in MOTOR MAKE -> "3RD PARTY"

    CAM Run Tracker:
    - All -> "CAM RENTAL"

    POG_CAM:
    - If JOB_TYPE is "RENTAL" -> "CAM RENTAL"
    - If JOB_TYPE is "DIRECTIONAL" -> "CAM DD"

    POG_MM:
    - All -> "TDI CONV"
    """
    if 'MOTOR_TYPE2' in df.columns:
        def determine_motor_type2(row):
            source = row['SOURCE']

            # Motor KPI logic
            if source == 'Motor_KPI':
                sn = str(row.get('SN', '')).upper() if pd.notna(row.get('SN')) else ''
                motor_make = str(row.get('MOTOR_MAKE', '')).upper() if pd.notna(row.get('MOTOR_MAKE')) else ''

                if 'MLA07' in sn:
                    return 'CAM DD'
                elif 'TDI' in motor_make and 'MLA07' not in sn:
                    return 'TDI CONV'
                else:
                    return '3RD PARTY'

            # CAM Run Tracker logic
            elif source == 'CAM_Run_Tracker':
                return 'CAM RENTAL'

            # POG_CAM logic
            elif source == 'POG_CAM_Usage':
                job_type = str(row.get('JOB_TYPE', '')).strip().upper() if pd.notna(row.get('JOB_TYPE')) else ''
                if 'RENTAL' in job_type:
                    return 'CAM RENTAL'
                elif 'DIRECTIONAL' in job_type:
                    return 'CAM DD'
                return None

            # POG_MM logic
            elif source == 'POG_MM_Usage':
                return 'TDI CONV'

            return None

        df['MOTOR_TYPE2'] = df.apply(determine_motor_type2, axis=1)
        print(f"  Populated MOTOR_TYPE2 column based on source-specific logic")

    return df

# ============================================================================
# STEP 8: Apply Lookups
# ============================================================================

def apply_basin_lookup(df, county_to_basin):
    """Apply basin lookup based on county"""
    if 'COUNTY' in df.columns and 'BASIN' in df.columns:
        def get_basin(county):
            if pd.isna(county):
                return None
            county_str = str(county).strip().upper()
            return county_to_basin.get(county_str, None)

        df['BASIN'] = df['COUNTY'].apply(get_basin)

    return df

def apply_formfam_lookup(df, formfam_df):
    """Apply formation family lookup"""
    if 'FORMATION' in df.columns and 'BASIN' in df.columns and 'FORM_FAM' in df.columns:
        # Create lookup dictionary
        formfam_dict = {}
        for _, row in formfam_df.iterrows():
            key = (str(row['Basin']).upper(), str(row['Keyword']).upper())
            formfam_dict[key] = row['Formation Family']

        def get_form_fam(row):
            if pd.isna(row['FORMATION']) or pd.isna(row['BASIN']):
                return None

            basin = str(row['BASIN']).upper()
            formation = str(row['FORMATION']).upper()

            # Look for keyword match in formation name
            for (lookup_basin, keyword), form_fam in formfam_dict.items():
                if lookup_basin == basin and keyword in formation:
                    return form_fam

            return None

        df['FORM_FAM'] = df.apply(get_form_fam, axis=1)

    return df

# ============================================================================
# STEP 5: Merge All Data
# ============================================================================

def merge_all_files():
    """Main function to merge all files"""

    print("="*80)
    print("EXCEL FILES MERGER - STARTING")
    print("="*80)

    # Step 1: Load mapping
    mappings, target_headers = load_mapping()

    # Step 2: Load lookup tables
    county_to_basin, formfam_df = load_lookup_tables()

    # Step 3: Read all source files
    dfs = []

    # Motor KPI
    df_motor = read_motor_kpi(FILES['Motor_KPI'], mappings['Motor_KPI'])
    df_motor = clean_county_names(df_motor, 'Motor_KPI')
    df_motor = standardize_operator_names(df_motor, 'Motor_KPI')
    dfs.append(df_motor)

    # CAM Run Tracker
    df_cam = read_cam_run_tracker(FILES['CAM_Run_Tracker'], mappings['CAM Run Tracker'])
    df_cam = clean_county_names(df_cam, 'CAM_Run_Tracker')
    df_cam = standardize_operator_names(df_cam, 'CAM_Run_Tracker')
    dfs.append(df_cam)

    # POG CAM Usage
    df_pog_cam = read_pog_cam_usage(FILES['POG_CAM_Usage'], mappings['POG_CAM_Usage'])
    df_pog_cam = clean_county_names(df_pog_cam, 'POG_CAM_Usage')
    df_pog_cam = standardize_operator_names(df_pog_cam, 'POG_CAM_Usage')
    dfs.append(df_pog_cam)

    # POG MM Usage
    df_pog_mm = read_pog_mm_usage(FILES['POG_MM_Usage'], mappings['POG_MM_Usage'])
    df_pog_mm = clean_county_names(df_pog_mm, 'POG_MM_Usage')
    df_pog_mm = standardize_operator_names(df_pog_mm, 'POG_MM_Usage')
    dfs.append(df_pog_mm)

    # Step 4: Concatenate all dataframes
    print("\n" + "="*80)
    print("MERGING DATA")
    print("="*80)

    df_merged = pd.concat(dfs, ignore_index=True, sort=False)
    print(f"\nTotal rows after merge: {len(df_merged)}")
    print(f"Total columns: {len(df_merged.columns)}")

    # Step 5: Ensure all target headers are present (add missing columns with NaN)
    for header in target_headers:
        if header not in df_merged.columns:
            df_merged[header] = np.nan

    # Step 6: Reorder columns to match target format
    df_merged = df_merged[target_headers]

    # Step 7: Apply lookups
    print("\nApplying lookup tables...")
    df_merged = apply_basin_lookup(df_merged, county_to_basin)
    df_merged = apply_formfam_lookup(df_merged, formfam_df)

    # Step 8: Format dates and create START_DATE/END_DATE
    print("\nFormatting dates and creating START_DATE/END_DATE...")
    df_merged = format_dates_and_datetimes(df_merged)

    # Step 9: Populate LOBE/STAGE and DDS columns
    print("\nPopulating LOBE/STAGE and DDS columns...")
    df_merged = populate_lobe_stage_and_dds(df_merged)

    # Step 10: Populate TOTAL HRS
    print("\nPopulating TOTAL HRS...")
    df_merged = populate_total_hrs(df_merged)

    # Step 11: Add UPDATE column
    print("\nAdding UPDATE column...")
    df_merged = add_update_column(df_merged)

    # Step 12: Populate MOTOR_TYPE2
    print("\nPopulating MOTOR_TYPE2...")
    df_merged = populate_motor_type2(df_merged)

    # Step 13: Export to Excel
    print("\n" + "="*80)
    print("EXPORTING RESULTS")
    print("="*80)

    print(f"\nWriting to: {OUTPUT_FILE}")
    df_merged.to_excel(OUTPUT_FILE, index=False, sheet_name='Merged Data')

    print("\n" + "="*80)
    print("MERGE COMPLETE!")
    print("="*80)
    print(f"\nOutput file: {OUTPUT_FILE}")
    print(f"Total rows: {len(df_merged)}")
    print(f"Total columns: {len(df_merged.columns)}")

    # Print summary by source
    print("\n" + "-"*80)
    print("DATA SUMMARY BY SOURCE")
    print("-"*80)
    source_counts = df_merged['SOURCE'].value_counts()
    for source, count in source_counts.items():
        print(f"  {source}: {count} rows")

    # Print some statistics
    print("\n" + "-"*80)
    print("COLUMN FILL STATISTICS (Top 20 most populated)")
    print("-"*80)

    fill_stats = []
    for col in df_merged.columns:
        non_null_count = df_merged[col].notna().sum()
        fill_pct = (non_null_count / len(df_merged)) * 100
        fill_stats.append({
            'Column': col,
            'Non-Null Count': non_null_count,
            'Fill %': fill_pct
        })

    fill_df = pd.DataFrame(fill_stats).sort_values('Fill %', ascending=False)
    print(fill_df.head(20).to_string(index=False))

    return df_merged

# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == "__main__":
    try:
        df_result = merge_all_files()
        print("\nScript completed successfully!")
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
