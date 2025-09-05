import re
import os
import pandas as pd
from tkinter import filedialog, Tk
from tkinter.messagebox import showinfo
from datetime import datetime
import platform
import subprocess
import warnings
from openpyxl.styles.stylesheet import Stylesheet
from xlsxwriter import Workbook

# Suppress openpyxl style warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def normalize_number(value):
    if not value or value.strip() == '':
        return ''
    value = value.strip().replace(",", "")
    if value.endswith('-'):
        value = '-' + value[:-1]
    try:
        return f"{float(value):,.3f}"
    except:
        return value


def open_file(filepath):
    """Open file with the default system application."""
    try:
        if platform.system() == 'Windows':
            os.startfile(filepath)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', filepath])
        else:  # Linux and others
            subprocess.call(['xdg-open', filepath])
    except Exception as e:
        messagebox.showerror('Error', f'Failed to open file:\n{e}')




def parse_date(date_str, expected_format=None):
    """Robust date parser that handles DD/MM/YYYY and YYYY/MM/DD formats."""
    date_str = str(date_str).strip()
    if not date_str or date_str.lower() in ['nan', 'nat', 'none']:
        return None

    # Replace different separators with '/'
    date_str = date_str.replace('-', '/').replace('.', '/')

    # Try DD/MM/YYYY first (common in banking)
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        return dt.date()
    except ValueError:
        pass

    # Try YYYY/MM/DD (SFMS format)
    try:
        dt = datetime.strptime(date_str, "%Y/%m/%d")
        return dt.date()
    except ValueError:
        pass

    # Fallback for other formats (optional)
    try:
        dt = datetime.strptime(date_str, "%m/%d/%Y")  # US format (last resort)
        return dt.date()
    except ValueError:
        pass

    print(f"⚠ Could not parse date: {date_str}")
    return None

def detect_type_from_name(name):
    name = str(name).lower().replace(" ", "").replace("-", "").replace("_", "").replace("(", "").replace(")", "")
    
    if 'neft' in name:
        if any(x in name for x in ['out', 'ouward', 'outward', 'ow', 'neftout', 'neftisooutgoing']):
            return 'NEFT OUTWARD'
        elif any(x in name for x in ['in', 'inward', 'settlement', 'iw', 'neftin', 'neftisoincoming']):
            return 'NEFT INWARD'
    elif 'rtgs' in name:
        if any(x in name for x in ['out', 'ouward', 'outward', 'ow', 'rtgsout', 'rtgsoutgoing']):
            return 'RTGS OUTWARD'
        elif any(x in name for x in ['in', 'inward', 'settlement', 'iw', 'rtgsin', 'rtgsincoming']):
            return 'RTGS INWARD'
    return None

def make_unique_columns(columns):
    seen = {}
    result = []
    for col in columns:
        if col not in seen:
            seen[col] = 1
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}.{seen[col]}")
    return result


def write_unmatched_enquiry_report(all_unmatched_cleaned, npci_matched, writer):
    """
    Write the unmatched enquiry report after removing records matched in NPCI.

    Parameters:
    - all_unmatched_cleaned: list of pd.DataFrame containing unmatched enquiry dataframes.
    - npci_matched: pd.DataFrame containing matched NPCI enquiry records.
    - writer: pd.ExcelWriter object to write the final output Excel.
    """

    # If no unmatched enquiry data, return early
    if not all_unmatched_cleaned:
        print("No unmatched enquiry records to write.")
        return

    # Combine all unmatched enquiry dataframes into one
    unmatched_final = pd.concat(all_unmatched_cleaned, ignore_index=True)

    if not npci_matched.empty:
        # Create composite key columns in both dfs to identify matches
        unmatched_final['match_key'] = (
            unmatched_final['TRACE NO'].astype(str) + '_' + unmatched_final['POST DATE'].astype(str)
        )
        npci_matched['match_key'] = (
            npci_matched['TRACE NO_ENQUIRY'].astype(str) + '_' + npci_matched['POST DATE_ENQUIRY'].astype(str)
        )

        # Remove records from unmatched_final that appear in npci_matched
        unmatched_final = unmatched_final[~unmatched_final['match_key'].isin(npci_matched['match_key'])].copy()

        # Drop helper column
        unmatched_final.drop(columns=['match_key'], inplace=True)

    # Columns to include in the final output - adjust as per your needs
    required_columns = [
        'FILE NAME', 'ACCOUNT NO', 'OPENING BALANCE', 'TOTAL DR/CR', 'CLOSING BALANCE',
        'BRANCH', 'TERM', 'USER', 'TXN CODE', 'POST DATE', 'TRACE NO',
        'AMOUNT DR', 'AMOUNT CR', 'BALANCE', 'FULL NARRATIVE', 'SOURCE FILE'
    ]

    # Add missing columns as empty strings if they do not exist
    for col in required_columns:
        if col not in unmatched_final.columns:
            unmatched_final[col] = ''

    # Write to Excel sheet
    unmatched_final[required_columns].to_excel(writer, sheet_name='Unmatched_EnquiryReport', index=False)

    print("✅ Unmatched_EnquiryReport sheet created with matched records removed.")


def match_data(cleaned_subset, input_df, transaction_type):
    id_col = None
    camt_mask = input_df.apply(lambda row: is_camt_record(row, transaction_type), axis=1)
    camt_df = input_df[camt_mask].copy()
    process_df = input_df[~camt_mask].copy()
    
    if transaction_type == 'RTGS INWARD':
        if 'Message Identifier' in input_df.columns:
            id_col = 'Message Identifier'
    else:
        for candidate in ['End To End Id', 'Transaction End Id']:
            if candidate in input_df.columns:
                id_col = candidate
                break

    if not id_col:
        print(f"❌ No valid Transaction ID column found for {transaction_type}.")
        return pd.DataFrame(), pd.DataFrame(), input_df, 0

    total_sfms = len(input_df)
    
    # Improved ID extraction with better alphanumeric handling
    input_df['TRANSACTION_ID'] = input_df[id_col].astype(str).str.extract(
        r'(?:/XUTR/)?([A-Za-z0-9]{12,})'  # Minimum 12 chars to avoid short matches
    )[0]
    
    invalid_tx_rows = input_df[input_df['TRANSACTION_ID'].isna()].copy()
    valid_input_df = input_df.dropna(subset=['TRANSACTION_ID'])

    input_df = valid_input_df.copy()
    cleaned_subset = cleaned_subset.copy()
    cleaned_subset['NARRATIVE_MATCH'] = cleaned_subset['FULL NARRATIVE'].astype(str).str.upper()

    matched = []
    unmatched_input = []
    matched_cleaned_ids = set()

    for _, in_row in input_df.iterrows():
        tx_id = in_row['TRANSACTION_ID'].strip().upper()
        if not tx_id:
            continue

        # Get last 6 alphanumeric characters
        suffix = tx_id[-6:] if len(tx_id) >= 6 else tx_id
        search_patterns = [suffix, tx_id]  # Try both suffix and full ID

        matched_this_row = False
        
        for pattern in search_patterns:
            matches = cleaned_subset[
                cleaned_subset['NARRATIVE_MATCH'].str.contains(pattern, regex=False, case=False, na=False)
            ]
            
            for _, clean_row in matches.iterrows():
                try:
                    # Robust date parsing
                    enquiry_date = parse_date(str(clean_row['POST DATE']))
                    sfms_date = parse_date(str(in_row.get('Settlement Date (YYYY/MM/DD)', '')))
                    if not enquiry_date or not sfms_date or enquiry_date != sfms_date:
                        continue

                    # SAFE AMOUNT CONVERSION
                    def safe_float_convert(x):
                        try:
                            return float(str(x).replace(",", "").strip())
                        except:
                            return 0.0
                            
                    sfms_amount = safe_float_convert(in_row.get('Amount (INR)', ''))
                    amt_dr = safe_float_convert(clean_row.get('AMOUNT DR', ''))
                    amt_cr = safe_float_convert(clean_row.get('AMOUNT CR', ''))
                    
                    # Amount comparison with tolerance
                    if not (abs(sfms_amount - amt_dr) < 0.01 or abs(sfms_amount - amt_cr) < 0.01):
                        continue

                    # Valid match found
                    joined = pd.concat([clean_row.add_suffix('_ENQUIRY'), in_row.add_suffix('_SFMS')])
                    matched.append(joined)
                    matched_cleaned_ids.add(clean_row.name)
                    matched_this_row = True
                    break
                    
                except Exception as e:
                    print(f"⚠ Error in validation: {e}")
                    continue
                    
            if matched_this_row:
                break

        if not matched_this_row:
            unmatched_input.append(in_row)

    matched_df = pd.DataFrame(matched)
    unmatched_cleaned = cleaned_subset[~cleaned_subset.index.isin(matched_cleaned_ids)].drop(columns=['NARRATIVE_MATCH'])
    unmatched_input_df = pd.DataFrame(unmatched_input)
    
    if not invalid_tx_rows.empty:
        unmatched_input_df = pd.concat([unmatched_input_df, invalid_tx_rows], ignore_index=True)

    return matched_df, unmatched_cleaned, unmatched_input_df, total_sfms

def process_sfms_file(file_path, transaction_type):
    """Read and process SFMS Excel file with support for multiple tables"""
    # Read the entire file without skipping rows initially
    raw = pd.read_excel(file_path, dtype=str, sheet_name='sheet1', header=None)
    
    # Define expected starting columns for each file type
    expected_start_columns = {
        'NEFT INWARD': ['SI No', 'Sequence No', 'Transaction Id'],
        'NEFT OUTWARD': ['Sl No', 'Sequence Number', 'Transaction Id'],
        'RTGS INWARD': ['SI No', 'Owner Address', 'Sequence No'],
        'RTGS OUTWARD': ['SI No', 'Owner Address', 'Sequence No']
    }
    
    target_columns = expected_start_columns.get(transaction_type, [])
    tables = []
    current_pos = 0
    
    # Process all tables in the file
    while current_pos < len(raw):
        # Find the next header row
        header_idx = None
        for i in range(current_pos, len(raw)):
            row = raw.iloc[i].astype(str).str.strip()
            if len(row) >= len(target_columns):
                if all(target_col.lower() in row[j].lower() 
                      for j, target_col in enumerate(target_columns)):
                    header_idx = i
                    break
        
        if header_idx is None:
            break  # No more tables found
            
        # Find the next "End of Records" after the header
        end_idx = None
        for i in range(header_idx + 1, len(raw)):
            row_str = ' '.join(raw.iloc[i].astype(str).str.lower())
            if 'end of records' in row_str:
                end_idx = i
                break
        
        if end_idx is None:
            end_idx = len(raw)
        
        # Extract this table
        headers = make_unique_columns(raw.iloc[header_idx].astype(str).str.strip().tolist())
        table_df = raw.iloc[header_idx + 1:end_idx].copy()
        table_df.columns = headers
        
        # Add to our tables list if it has data
        if len(table_df) > 0:
            tables.append(table_df)
        
        # Move position to after this table
        current_pos = end_idx + 1
    
    if not tables:
        print(f"❌ No valid tables found in {os.path.basename(file_path)}")
        return None
    
    # Combine all tables into one dataframe
    combined_df = pd.concat(tables, ignore_index=True)
    return combined_df

def is_camt_record(row, transaction_type):
    """Check if a record is a CAMT record"""
    if 'RTGS' in transaction_type:
        return 'camt.059.001.04' in str(row.get('Message Type', '')).lower()
    else:  # NEFT
        return any('camt' in str(val).lower() for val in row)

def process_transaction_enquiry_files(file_paths):
    all_records = []
    columns = [
        "FILE NAME", "ACCOUNT NO", "OPENING BALANCE", "TOTAL DR/CR", "CLOSING BALANCE",
        "BRANCH", "TERM", "USER", "TXN CODE", "POST DATE", "TRACE NO",
        "AMOUNT DR", "AMOUNT CR", "BALANCE", "FULL NARRATIVE"
    ]
    previous_closing_balance = ""

    for file_path in file_paths:
        custom_label = ""
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

        for line in lines:
            if "TRANSACTION ENQUIRY REPORT FOR ACCOUNT NO" in line:
                match = re.search(r"ACCOUNT NO\s*:\s*\d+\s*-\s*(.*?)\s+FROM DATE", line)
                if match:
                    custom_label = match.group(1).strip()
                break

        full_file_name = custom_label if custom_label else os.path.basename(file_path)

        current_block = {
            "account": "",
            "opening_balance": previous_closing_balance,
            "total_dr_cr": "",
            "closing_balance": "",
            "transactions": []
        }

        for raw_line in lines:
            line = raw_line.strip()

            if "ACCOUNT NO" in line:
                if current_block["transactions"]:
                    for transaction in current_block["transactions"]:
                        transaction[3] = current_block["total_dr_cr"]
                        transaction[4] = current_block["closing_balance"]
                        all_records.append(transaction)
                    previous_closing_balance = current_block["closing_balance"]

                account_match = re.search(r"ACCOUNT NO\s*:\s*(\d+)", line)
                current_block = {
                    "account": account_match.group(1) if account_match else "",
                    "opening_balance": previous_closing_balance,
                    "total_dr_cr": "",
                    "closing_balance": "",
                    "transactions": []
                }

            elif "OPENING BALANCE" in line:
                open_match = re.search(r"OPENING BALANCE\s*:\s*([\d,\.-]+)\s*(CR|DR)?", line, re.IGNORECASE)
                if open_match:
                    current_block["opening_balance"] = f"{normalize_number(open_match.group(1))} {(open_match.group(2) or '').strip()}"

            elif re.match(r"^\d{5}", raw_line) and len(raw_line) > 100:
                try:
                    full_narr = raw_line[138:].strip()
                    transaction = [
                        full_file_name,
                        current_block["account"],
                        current_block["opening_balance"],
                        "", "",
                        raw_line[9:14].strip(),
                        raw_line[15:20].strip(),
                        raw_line[21:31].strip(),
                        raw_line[32:40].strip(),
                        raw_line[41:53].strip(),
                        raw_line[54:63].strip(),
                        normalize_number(raw_line[64:87].strip()),
                        normalize_number(raw_line[88:110].strip()),
                        normalize_number(raw_line[111:138].strip()),
                        full_narr
                    ]
                    current_block["transactions"].append(transaction)
                except Exception as e:
                    print(f"⚠ Error processing transaction line:\n{raw_line}\nError: {str(e)}\n")

            elif "TOTAL DR/CR" in line or "TOTAL   DR/CR" in line:
                total_match = re.search(r"TOTAL\s+DR/CR\s*:\s*([\d,\.-]+)", line, re.IGNORECASE)
                if total_match:
                    current_block["total_dr_cr"] = normalize_number(total_match.group(1))

            elif "CARRIED FORWARD" in line:
                cf_match = re.search(r"CARRIED FORWARD\s*:\s*([\d,\.-]+)\s*(CR|DR)?", line, re.IGNORECASE)
                if cf_match:
                    current_block["closing_balance"] = f"{normalize_number(cf_match.group(1))} {(cf_match.group(2) or '').strip()}"

        if current_block["transactions"]:
            for transaction in current_block["transactions"]:
                transaction[3] = current_block["total_dr_cr"]
                transaction[4] = current_block["closing_balance"]
                all_records.append(transaction)
            previous_closing_balance = current_block["closing_balance"]

    return pd.DataFrame(all_records, columns=columns)

def process_npci_files(file_paths):
    """Process NPCI files with guaranteed column preservation and file type detection"""
    all_npci = []
    
    for file_path in file_paths:
        try:
            # Read CSV with proper encoding
            df = pd.read_csv(file_path, encoding='utf-8-sig', skiprows=1, dtype=str)
            
            # Clean column names (keep originals)
            df.columns = [col.strip('"').strip() for col in df.columns]
            
            # Verify critical columns exist
            required_cols = ['No.', 'Reference', 'Business Date', 'Amount']
            if not all(col in df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in df.columns]
                print(f"❌ Missing in {os.path.basename(file_path)}: {missing}")
                continue
                
            # Add source tracking based on file name
            file_name = os.path.basename(file_path).upper()
            if 'DR' in file_name:
                df['ORIGINAL_FILE'] = 'NPCI_DR'
            elif 'CR' in file_name:
                df['ORIGINAL_FILE'] = 'NPCI_CR'
            else:
                # Default to DR if cannot determine from filename
                df['ORIGINAL_FILE'] = 'NPCI_DR'
                print(f"⚠ Could not determine file type from name: {file_name}. Defaulting to DR")
            
            # Reset index to avoid duplication issues
            df = df.reset_index(drop=True)
            all_npci.append(df)
            
        except Exception as e:
            print(f"⚠ Error processing {os.path.basename(file_path)}: {str(e)}")
    
    return pd.concat(all_npci, ignore_index=True) if all_npci else pd.DataFrame()


def match_npci_with_cbs(npci_df, cbs_df):
    """Robust matching with duplicate prevention and proper formatting"""
    try:
        # Filter CBS data to only include RTGS transactions
        rtgs_cbs = cbs_df[
            cbs_df['FILE TYPE'].isin(['RTGS INWARD', 'RTGS OUTWARD'])
        ].copy()
        
        # Reset indices to prevent duplication issues
        npci_df = npci_df.reset_index(drop=True)
        rtgs_cbs = rtgs_cbs.reset_index(drop=True)
        
        # Date standardization
        npci_df['MATCH_DATE'] = npci_df['Business Date'].apply(
            lambda x: parse_date(str(x).replace('-', '/')))
        rtgs_cbs['MATCH_DATE'] = rtgs_cbs['POST DATE'].apply(
            lambda x: parse_date(str(x).replace('/', '-')))
        
        # Rest of the matching logic remains the same...        
        matched = []
        unmatched_npci = []
        
        for npci_idx, npci_row in npci_df.iterrows():
            ref = str(npci_row['Reference']).strip()
            amount = abs(float(str(npci_row['Amount']).replace(',', '')))
            npci_date = npci_row['MATCH_DATE']
            
            if pd.isna(npci_date):
                unmatched_npci.append(npci_row)
                continue
                
            match_found = False
            for cbs_idx, cbs_row in cbs_df.iterrows():
                if npci_date != cbs_row['MATCH_DATE']:
                    continue
                    
                cbs_amount = abs(float(str(cbs_row['AMOUNT DR'] or cbs_row['AMOUNT CR']).replace(',', '')))
                if abs(amount - cbs_amount) > 0.01:
                    continue
                    
                if ref in str(cbs_row['FULL NARRATIVE']):
                    # Format the matched record with CBS columns first, then NPCI columns
                    matched_record = {
                        # CBS columns
                        'FILE NAME_ENQUIRY': cbs_row['FILE NAME'],
                        'ACCOUNT NO_ENQUIRY': cbs_row['ACCOUNT NO'],
                        'OPENING BALANCE_ENQUIRY': cbs_row['OPENING BALANCE'],
                        'TOTAL DR/CR_ENQUIRY': cbs_row['TOTAL DR/CR'],
                        'CLOSING BALANCE_ENQUIRY': cbs_row['CLOSING BALANCE'],
                        'BRANCH_ENQUIRY': cbs_row['BRANCH'],
                        'TERM_ENQUIRY': cbs_row['TERM'],
                        'USER_ENQUIRY': cbs_row['USER'],
                        'TXN CODE_ENQUIRY': cbs_row['TXN CODE'],
                        'POST DATE_ENQUIRY': cbs_row['POST DATE'],
                        'TRACE NO_ENQUIRY': cbs_row['TRACE NO'],
                        'AMOUNT DR_ENQUIRY': cbs_row['AMOUNT DR'],
                        'AMOUNT CR_ENQUIRY': cbs_row['AMOUNT CR'],
                        'BALANCE_ENQUIRY': cbs_row['BALANCE'],
                        'FULL NARRATIVE_ENQUIRY': cbs_row['FULL NARRATIVE'],
                        
                        # NPCI columns
                        'No._NPCI': npci_row['No.'],
                        'Reference_NPCI': npci_row['Reference'],
                        'Business Date_NPCI': npci_row['Business Date'],
                        'Amount_NPCI': npci_row['Amount'],
                        'ORIGINAL_FILE_NPCI': npci_row['ORIGINAL_FILE']
                    }
                    matched.append(matched_record)
                    match_found = True
                    break
                    
            if not match_found:
                unmatched_npci.append(npci_row)
                
        return pd.DataFrame(matched), pd.DataFrame(unmatched_npci)
        
    except Exception as e:
        print(f"⚠ Matching error: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()


def safe_float(val):
    # Convert to string first
    s = str(val).replace(',', '').strip()
    # If empty or not a valid number, return 0
    return float(s) if s else 0

def main():
    # Initialize Tkinter root window (hidden)
    root = Tk()
    root.withdraw()

    # File selection dialogs
    enquiry_files = filedialog.askopenfilenames(
        title="Select TRANSACTION ENQUIRY REPORT files",
        filetypes=[("Text files", ".txt .prn *.prt"), ("All files", ".")]
    )
    sfms_files = filedialog.askopenfilenames(
        title="Select SFMS NEFT/RTGS Excel Files",
        filetypes=[("Excel files", "*.xlsx")]

    )
    npci_files = filedialog.askopenfilenames(
        title="Select NPCI CSV Files",
        filetypes=[("CSV files", "*.csv")]
    )
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        title="Save Final Output Excel As",
        filetypes=[("Excel files", "*.xlsx")]
    )

    # Validate file selections
    if not enquiry_files or not output_file:
        print("File selection was canceled.")
        return

    # ======================
    # 1. PROCESS ENQUIRY FILES
    # ======================
    print("\n" + "="*50)
    print("Processing transaction enquiry files...")
    print("="*50)
    
    cleaned_df = process_transaction_enquiry_files(enquiry_files)
    cleaned_df['FILE TYPE'] = cleaned_df['FILE NAME'].apply(detect_type_from_name)
    print(f"Processed {len(cleaned_df)} enquiry records")

    # ======================
    # 2. PROCESS NPCI FILES
    # ======================
    npci_counts = {
        'TOTAL_DR': 0,
        'TOTAL_CR': 0,
        'UNMATCHED_DR': 0,
        'UNMATCHED_CR': 0
    }
    npci_matched = pd.DataFrame()
    npci_unmatched = pd.DataFrame()

    # Check if RTGS files are present for NPCI processing
    rtgs_files_present = any(
        detect_type_from_name(os.path.basename(f)) in ['RTGS INWARD', 'RTGS OUTWARD']
        for f in sfms_files
    )

    if npci_files :
        print("\n" + "="*50)
        print("Processing NPCI files...")
        print("="*50)
        
        npci_df = process_npci_files(npci_files)
        
        if not npci_df.empty:
            # Ensure ORIGINAL_FILE column exists
            if 'ORIGINAL_FILE' not in npci_df.columns:
                print("⚠ ORIGINAL_FILE column not found - defaulting to DR")
                npci_df['ORIGINAL_FILE'] = 'NPCI_DR'
            
            # Calculate counts
            npci_counts['TOTAL_DR'] = len(npci_df[npci_df['ORIGINAL_FILE'] == 'NPCI_DR'])
            npci_counts['TOTAL_CR'] = len(npci_df[npci_df['ORIGINAL_FILE'] == 'NPCI_CR'])
            print(f"Found {npci_counts['TOTAL_DR']} DR records and {npci_counts['TOTAL_CR']} CR records")

            # Filter CBS data for RTGS only
            rtgs_cbs = cleaned_df[
                cleaned_df['FILE TYPE'].isin(['RTGS INWARD', 'RTGS OUTWARD'])
            ].copy()
            
            # Enhanced matching function
            def match_npci_with_cbs(npci_df, cbs_df):
                try:
                    # Reset indices
                    npci_df = npci_df.reset_index(drop=True)
                    cbs_df = cbs_df.reset_index(drop=True)
                    
                    # Convert dates with multiple format support
                    date_formats = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d']
                    
                    npci_df['MATCH_DATE'] = pd.to_datetime(
                        npci_df['Business Date'],
                        dayfirst=True,
                        errors='coerce'
                    ).dt.date
                    
                    cbs_df['MATCH_DATE'] = pd.to_datetime(
                        cbs_df['POST DATE'],
                        dayfirst=True,
                        errors='coerce'
                    ).dt.date
                    
                    # Drop rows with invalid dates
                    npci_df = npci_df.dropna(subset=['MATCH_DATE'])
                    cbs_df = cbs_df.dropna(subset=['MATCH_DATE'])
                    
                    matched = []
                    unmatched_npci = []
                    
                    for npci_idx, npci_row in npci_df.iterrows():
                        ref = str(npci_row['Reference']).strip()
                        amount = abs(float(str(npci_row['Amount']).replace(',', '')))
                        npci_date = npci_row['MATCH_DATE']
                        
                        match_found = False
                        for cbs_idx, cbs_row in cbs_df.iterrows():
                            # Skip if dates don't match
                            if npci_date != cbs_row['MATCH_DATE']:
                                continue
                            
                            # Get amount from CBS (try DR then CR)
                            cbs_amount = 0.0
                            if 'AMOUNT DR' in cbs_row and str(cbs_row['AMOUNT DR']) not in ['', 'nan']:
                                cbs_amount = abs(float(str(cbs_row['AMOUNT DR']).replace(',', '')))
                            elif 'AMOUNT CR' in cbs_row and str(cbs_row['AMOUNT CR']) not in ['', 'nan']:
                                cbs_amount = abs(float(str(cbs_row['AMOUNT CR']).replace(',', '')))
                            
                            # Amount comparison with 0.01 tolerance
                            if abs(amount - cbs_amount) > 0.01:
                                continue
                            
                            # Reference check in narrative
                            narrative = str(cbs_row.get('FULL NARRATIVE', ''))
                            if ref and ref in narrative:
                                # Create matched record
                                matched_record = {
                                    **{f'{k}_ENQUIRY': v for k, v in cbs_row.items()},
                                    **{f'{k}_NPCI': v for k, v in npci_row.items()}
                                }
                                matched.append(matched_record)
                                match_found = True
                                break
                        
                        if not match_found:
                            unmatched_npci.append(npci_row)
                    
                    return pd.DataFrame(matched), pd.DataFrame(unmatched_npci)
                
                except Exception as e:
                    print(f"⚠ Error in NPCI matching: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    return pd.DataFrame(), pd.DataFrame()

            # Perform matching
            npci_matched, npci_unmatched = match_npci_with_cbs(npci_df, rtgs_cbs)
            print(f"Matched {len(npci_matched)} NPCI records")

            # Update unmatched counts
            if not npci_unmatched.empty:
                npci_counts['UNMATCHED_DR'] = len(npci_unmatched[npci_unmatched['ORIGINAL_FILE'] == 'NPCI_DR'])
                npci_counts['UNMATCHED_CR'] = len(npci_unmatched[npci_unmatched['ORIGINAL_FILE'] == 'NPCI_CR'])
                print(f"Unmatched NPCI: {npci_counts['UNMATCHED_DR']} DR, {npci_counts['UNMATCHED_CR']} CR")

            # No changes here to cleaned_df directly for now (we'll remove matched records only for Unmatched_EnquiryReport sheet below)
    
    elif npci_files:
        print("\n⚠ NPCI files provided but no RTGS files found - skipping NPCI processing")

    # ======================
    # 3. PROCESS SFMS FILES
    # ======================
    print("\n" + "="*50)
    print("Processing SFMS files...")
    print("="*50)
    
    all_matched = []
    all_unmatched_cleaned = []
    all_unmatched_input = []
    all_camt = []
    summary = []

    # (Your existing SFMS processing code here remains unchanged...)

    # =====
    # (Assuming the loop for SFMS files populates all_unmatched_cleaned, all_matched etc. as in your code)
    # =====


    # Original unmatched headers dictionary
    unmatched_input_headers = {
        'NEFT INWARD': ['SI No', 'Sequence No', 'Transaction Id', 'Return Id', 'End To End Id', 
                       'Sender IFSC', 'Sender Account Type', 'Sender Account Number', 'Sender Name', 
                       'Amount (INR)', 'Settlement Date (YYYY/MM/DD)', 'Creation Date (YYYY/MM/DD)', 
                       'Batch Number', 'Fresh / Return / Rejected', 'Beneficiary IFSC', 
                       'Transaction Status', 'Beneficiary Account Type', 'Beneficiary Account Number', 
                       'Beneficiary Name', 'Return Reject Code', 'Return/Reject\n Reason', 
                       'Unstructured', 'Instruction Info'],
        'NEFT OUTWARD': ['Sl No', 'Sequence Number', 'Transaction Id', 'End To End Id', '', '', 
                         'Amount (INR)', 'Settlement Date (YYYY/MM/DD)', 'Batch Number', 
                         'Sender IFSC', 'Sender Account Type', 'Sender Account Number', 
                         'Sender Account Name', 'Beneficiary IFSC', 'Beneficiary Account Type', 
                         'Beneficiary Account Number', 'Beneficiary Account Name', 
                         'Transaction Status', 'Instruction Info', 'Remittance Info', '', ''],
        'RTGS INWARD': ['SI No', 'Owner Address', 'Sequence No', 'Message Type', 'Amount (INR)', 
                       'Message Identifier', 'Transaction Id', 'Transaction End Id', 'Debitor Agent', 
                       'Debitor Account', 'Debitor Name', 'Debitor FINInst Name', 'DCCB', 
                       'Creditor Agent', 'Creditor Account', 'Creditor Name', 'Creditor FinInst Name', 
                       'Service Level Proprietory', 'Account Type', 'Institution Information', 
                       'Remittance Information', 'Transaction Ret Code', 'Transaction Ret Reason', 
                       'Value Date (YYYY/MM/DD)', 'Settlement Date (YYYY/MM/DD)', 
                       'Received Time(HH:MM:SS)', 'Transaction Status'],
        'RTGS OUTWARD': ['SI No', 'Owner Address', 'Sequence No', 'Message Type', 'Amount (INR)', 
                        'Message Identifier', 'Rel Message Identifier', 'Xml Utr Number', 
                        'Transaction Id', 'Rel Transaction Id', 'Transaction End Id', '', '', 
                        'Rel Transaction End Id', 'Debitor Agent', 'Debitor Account', 'Debitor Name', 
                        'Debitor FINInst Name', 'Creditor Agent', 'Creditor Account', 'Creditor Name', 
                        'Creditor FinInst Name', 'Service Level Proprietory', 'Account Type', 
                        'Institution Information', 'Remittance Information', 'Transaction Ret Code', 
                        'Transaction Ret Reason', 'Value Date (YYYY/MM/DD)', 'Settlement Date (YYYY/MM/DD)', 
                        'Settlement Time (HH:MM:SS)', 'Transaction Status', 'Failure Code', 'Failure Reason']
    }


    if sfms_files:
        for file_idx, file in enumerate(sfms_files, 1):
            file_name = os.path.basename(file)
            print(f"\n[{file_idx}/{len(sfms_files)}] Processing {file_name}...")
            
            try:
                # Detect transaction type
                input_type = detect_type_from_name(file_name)
                if not input_type:
                    print(f"⚠ Could not determine type for {file_name}")
                    continue
                
                # Process SFMS file
                input_df = process_sfms_file(file, input_type)
                if input_df is None:
                    print(f"⚠ No valid data in {file_name}")
                    continue
                
                # Separate CAMT records
                camt_mask = input_df.apply(lambda row: is_camt_record(row, input_type), axis=1)
                camt_df = input_df[camt_mask].copy()
                non_camt_df = input_df[~camt_mask].copy()
                
                # Extract transaction IDs based on type
                if 'RTGS' in input_type:
                    id_col = 'Message Identifier' if 'Message Identifier' in non_camt_df.columns else None
                else:  # NEFT
                    id_col = 'End To End Id' if 'End To End Id' in non_camt_df.columns else None
                
                if not id_col:
                    print(f"❌ No ID column found in {file_name}")
                    continue
                    
                non_camt_df['TRANSACTION_ID'] = non_camt_df[id_col].astype(str).str.extract(r'([A-Z0-9]{10,})')[0]
                
                # Get matching records from enquiry data
                subset = cleaned_df[cleaned_df['FILE TYPE'] == input_type]
                total_enquiry_records = len(subset)
                
                if subset.empty:
                    print(f"⚠ No matching enquiry records for {input_type}")
                    continue

                # Perform matching
                matched, unmatched_cleaned, unmatched_input, total_sfms = match_data(
                    subset,
                    non_camt_df,
                    input_type
                )
                print(f"Matched {len(matched)}/{total_sfms} records")
                
                # Track CAMT records
                if not camt_df.empty:
                    camt_df['SOURCE FILE'] = file_name
                    all_camt.append(camt_df)
                    print(f"Found {len(camt_df)} CAMT records")

                # Add source file info
                matched['SOURCE FILE'] = file_name
                unmatched_cleaned['SOURCE FILE'] = file_name
                unmatched_input['SOURCE FILE'] = file_name

                # Store results
                all_matched.append(matched)
                all_unmatched_cleaned.append(unmatched_cleaned)
                all_unmatched_input.append(unmatched_input)

                # Create summary entry

                npc_files_selected = (
                    bool(npci_counts) and
                    ('TOTAL_CR' in npci_counts or 'TOTAL_DR' in npci_counts)
                )


    # Build the base summary entry
                summary_entry = {
                    'SFMS File': file_name,
                    'Transaction Type': input_type,
                    'Total Records in SFMS': total_sfms + len(camt_df),
                    'Total Records in TRANSACTION ENQUIRY REPORT': total_enquiry_records,
                    'Matched Records': len(matched),
                    'Unmatched in TRANSACTION ENQUIRY REPORT': len(unmatched_cleaned),
                    'Unmatched in SFMS': len(unmatched_input),
                    'CAMT Records': len(camt_df),
                }

    # Add NPCI-related info only if selected
                if npci_files:
                    if input_type == 'RTGS INWARD':
                        total_npci = npci_counts.get('TOTAL_CR', 0)
                        unmatched_npci = npci_counts.get('UNMATCHED_CR', 0)
                    elif input_type == 'RTGS OUTWARD':
                        total_npci = npci_counts.get('TOTAL_DR', 0)
                        unmatched_npci = npci_counts.get('UNMATCHED_DR', 0)
                    else:
                        total_npci = unmatched_npci = 0

                    matched_npci = total_npci - unmatched_npci

                    summary_entry.update({
                        'Total records in NPCI': total_npci,
                        'Unmatched in NPCI': unmatched_npci,
                        'Matched Records': summary_entry['Matched Records'] + matched_npci,
                        'Unmatched in TRANSACTION ENQUIRY REPORT': max(
                            0,
                            summary_entry['Unmatched in TRANSACTION ENQUIRY REPORT'] - matched_npci
                        )
                    })

                summary.append(summary_entry)

                
            except Exception as e:
                print(f"⚠ Error processing {file_name}: {str(e)}")
                import traceback
                traceback.print_exc()
        print("\n" + "="*50)
        print("Generating output file...")
        print("="*50)
    
    elif (npci_files):
        for file_idx, file in enumerate(enquiry_files, 1):
            file_name = os.path.basename(file)
            input_type = detect_type_from_name(file_name)
            subset = cleaned_df[cleaned_df['FILE TYPE'] == input_type]
            total_enquiry_records = len(subset)
            npc_files_selected = (
                bool(npci_counts) and
                ('TOTAL_CR' in npci_counts or 'TOTAL_DR' in npci_counts)
            )

                  

                # Add source file info
           
    # Build the base summary entry
            summary_entry = {
                'Transaction Type': input_type,
                'Total Records in TRANSACTION ENQUIRY REPORT': total_enquiry_records,
            }

    # Add NPCI-related info only if selected
            if npci_files:
                if input_type == 'RTGS INWARD':
                    total_npci = npci_counts.get('TOTAL_CR', 0)
                    unmatched_npci = npci_counts.get('UNMATCHED_CR', 0)
                elif input_type == 'RTGS OUTWARD':
                    total_npci = npci_counts.get('TOTAL_DR', 0)
                    unmatched_npci = npci_counts.get('UNMATCHED_DR', 0)
                else:
                    total_npci = unmatched_npci = 0

                matched_npci = total_npci - unmatched_npci

                summary_entry.update({
                    'Total records in NPCI': total_npci,
                    'Unmatched in NPCI': unmatched_npci,
                    'Matched Records':  matched_npci,
                    'Unmatched in TRANSACTION ENQUIRY REPORT': max(
                        0,
                        total_enquiry_records - matched_npci
                    )
                })

            summary.append(summary_entry)

    else:
        summary_entry = {
            'Transaction Type': input_type,
            'Total Records in TRANSACTION ENQUIRY REPORT': total_enquiry_records,
            'Matched Records': len(matched),
            'Unmatched in TRANSACTION ENQUIRY REPORT': len(unmatched_cleaned),
        }



    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # 1. Matched Records Sheet
            if all_matched or not npci_matched.empty:
                matched_dfs = []
                if all_matched:
                    matched_dfs.append(pd.concat(all_matched, ignore_index=True))
                if not npci_matched.empty:
                    matched_dfs.append(npci_matched)
                
                matched_final = pd.concat(matched_dfs, ignore_index=True)
                matched_final.to_excel(writer, sheet_name='Matched', index=False)
                print("✅ Created Matched sheet")
            
            # 2. Unmatched Enquiry Report
            # Build Unmatched_EnquiryReport even if only CBS + NPCI is provided
            if all_unmatched_cleaned or not sfms_files:
                if all_unmatched_cleaned:
                    unmatched_cleaned_final = pd.concat(all_unmatched_cleaned, ignore_index=True)
                else:
                    # Use full CBS enquiry data for RTGS types only
                    unmatched_cleaned_final = cleaned_df.copy()                
                # Build match keys
                unmatched_cleaned_final['match_key'] = unmatched_cleaned_final.apply(
                    lambda row: (
                        row.get('FILE NAME', ''),
                        row.get('ACCOUNT NO', ''),
                        pd.to_datetime(row.get('POST DATE', '')).date() if pd.notna(row.get('POST DATE', None)) else None,
                        safe_float(row.get('AMOUNT DR', 0)),
                        safe_float(row.get('AMOUNT CR', 0))
                    ),
                    axis=1
                )

                npci_matched['match_key'] = npci_matched.apply(
                    lambda row: (
                        row.get('FILE NAME_ENQUIRY', ''),
                        row.get('ACCOUNT NO_ENQUIRY', ''),
                        pd.to_datetime(row.get('POST DATE_ENQUIRY', '')).date() if pd.notna(row.get('POST DATE_ENQUIRY', None)) else None,
                        safe_float(row.get('AMOUNT DR_ENQUIRY', 0)),
                        safe_float(row.get('AMOUNT CR_ENQUIRY', 0))
                    ),
                    axis=1
                )

                # Filter CBS records that were not matched to NPCI
                matched_keys = set(npci_matched['match_key'])
                filtered_unmatched_cleaned = unmatched_cleaned_final[
                    ~unmatched_cleaned_final['match_key'].isin(matched_keys)
                ].drop(columns=['match_key'])

                filtered_unmatched_cleaned.to_excel(writer, sheet_name='Unmatched_EnquiryReport', index=False)
                print("✅ Created Unmatched_EnquiryReport sheet (filtered matched NPCI records)")

            
            # 3. Unmatched NPCI
            if not npci_unmatched.empty:
                npci_cols = ['No.', 'Reference', 'Business Date', 'Amount', 'ORIGINAL_FILE']
                available_cols = [col for col in npci_cols if col in npci_unmatched.columns]
                npci_unmatched[available_cols].to_excel(
                    writer, sheet_name='Unmatched_NPCI', index=False)
                print("✅ Created Unmatched_NPCI sheet")
            
            # 4. Unmatched SFMS (separate sheets by type with proper headers)
                        # 4. Unmatched SFMS (separate sheets by type with proper headers)
            if all_unmatched_input:
                unmatched_input_combined = pd.concat(all_unmatched_input, ignore_index=True)
                
                # Define headers for each type
                unmatched_input_headers = {
                    'NEFT INWARD': ['SI No', 'Sequence No', 'Transaction Id', 'Return Id', 'End To End Id', 
                                   'Sender IFSC', 'Sender Account Type', 'Sender Account Number', 'Sender Name', 
                                   'Amount (INR)', 'Settlement Date (YYYY/MM/DD)', 'Creation Date (YYYY/MM/DD)', 
                                   'Batch Number', 'Fresh / Return / Rejected', 'Beneficiary IFSC', 
                                   'Transaction Status', 'Beneficiary Account Type', 'Beneficiary Account Number', 
                                   'Beneficiary Name', 'Return Reject Code', 'Return/Reject\n Reason', 
                                   'Unstructured', 'Instruction Info'],
                    'NEFT OUTWARD': ['Sl No', 'Sequence Number', 'Transaction Id', 'End To End Id', '', '', 
                                     'Amount (INR)', 'Settlement Date (YYYY/MM/DD)', 'Batch Number', 
                                     'Sender IFSC', 'Sender Account Type', 'Sender Account Number', 
                                     'Sender Account Name', 'Beneficiary IFSC', 'Beneficiary Account Type', 
                                     'Beneficiary Account Number', 'Beneficiary Account Name', 
                                     'Transaction Status', 'Instruction Info', 'Remittance Info', '', ''],
                    'RTGS INWARD': ['SI No', 'Owner Address', 'Sequence No', 'Message Type', 'Amount (INR)', 
                                   'Message Identifier', 'Transaction Id', 'Transaction End Id', 'Debitor Agent', 
                                   'Debitor Account', 'Debitor Name', 'Debitor FINInst Name', 'DCCB', 
                                   'Creditor Agent', 'Creditor Account', 'Creditor Name', 'Creditor FinInst Name', 
                                   'Service Level Proprietory', 'Account Type', 'Institution Information', 
                                   'Remittance Information', 'Transaction Ret Code', 'Transaction Ret Reason', 
                                   'Value Date (YYYY/MM/DD)', 'Settlement Date (YYYY/MM/DD)', 
                                   'Received Time(HH:MM:SS)', 'Transaction Status'],
                    'RTGS OUTWARD': ['SI No', 'Owner Address', 'Sequence No', 'Message Type', 'Amount (INR)', 
                                    'Message Identifier', 'Rel Message Identifier', 'Xml Utr Number', 
                                    'Transaction Id', 'Rel Transaction Id', 'Transaction End Id', '', '', 
                                    'Rel Transaction End Id', 'Debitor Agent', 'Debitor Account', 'Debitor Name', 
                                    'Debitor FINInst Name', 'Creditor Agent', 'Creditor Account', 'Creditor Name', 
                                    'Creditor FinInst Name', 'Service Level Proprietory', 'Account Type', 
                                    'Institution Information', 'Remittance Information', 'Transaction Ret Code', 
                                    'Transaction Ret Reason', 'Value Date (YYYY/MM/DD)', 'Settlement Date (YYYY/MM/DD)', 
                                    'Settlement Time (HH:MM:SS)', 'Transaction Status', 'Failure Code', 'Failure Reason']
                }
                
                for trans_type in ['NEFT INWARD', 'NEFT OUTWARD', 'RTGS INWARD', 'RTGS OUTWARD']:
                    # Filter by transaction type using .copy() to avoid SettingWithCopyWarning
                    subset = unmatched_input_combined.loc[
                        unmatched_input_combined['SOURCE FILE'].apply(
                            lambda x: detect_type_from_name(x) == trans_type)
                    ].copy()
                    
                    if not subset.empty:
                        sheet_name = f'Unmatched_SFMS_{trans_type.replace(" ", "")}'
                        
                        # Get the expected headers for this type
                        expected_headers = unmatched_input_headers.get(trans_type, [])
                        
                        # Add missing columns with empty values using .loc
                        for col in expected_headers:
                            if col not in subset.columns and col.strip():  # Skip empty column names
                                subset.loc[:, col] = ''
                        
                        # Reorder columns to match expected headers
                        ordered_cols = [col for col in expected_headers if col in subset.columns]
                        subset[ordered_cols].to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"✅ Created {sheet_name} sheet with proper headers")
            # 5. CAMT Records (separate sheets by type)
            if all_camt:
                camt_combined = pd.concat(all_camt, ignore_index=True)
                for trans_type in ['NEFT INWARD', 'NEFT OUTWARD', 'RTGS INWARD', 'RTGS OUTWARD']:
                    subset = camt_combined[
                        camt_combined['SOURCE FILE'].apply(
                            lambda x: detect_type_from_name(x) == trans_type)
                    ]
                    if not subset.empty:
                        sheet_name = f'CAMT_{trans_type.replace(" ", "")}'
                        subset.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"✅ Created {sheet_name} sheet")
            
            # 6. Summary Sheet
            if summary:
                summary_df = pd.DataFrame(summary)
                # Reorder columns
                column_order = [
                    'SFMS File', 'Transaction Type',
                    'Total Records in SFMS', 'Total Records in TRANSACTION ENQUIRY REPORT',
                    'Matched Records', 'Unmatched in TRANSACTION ENQUIRY REPORT',
                    'Unmatched in SFMS', 'Total records in NPCI', 'Unmatched in NPCI',
                    'CAMT Records'
                ]
                summary_df = summary_df.reindex(columns=[col for col in column_order if col in summary_df.columns])
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Add auto-filter
                worksheet = writer.sheets['Summary']
                worksheet.autofilter(0, 0, len(summary_df), len(summary_df.columns) - 1)
                print("✅ Created Summary sheet with filters")
                        
        print(f"\n🎉 Successfully created output file: {output_file}")
        showinfo("Success", f"Processing complete!\nOutput saved to:\n{output_file}")
        open_file(output_file)

    except Exception as e:
        error_msg = f"Error generating output file: {str(e)}"
        print(f"⚠ {error_msg}")
        showinfo("Error", error_msg)
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()