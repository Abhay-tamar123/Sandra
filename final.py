import pandas as pd
import re
import os
import gc
import win32com.client as win32
import tempfile
import pythoncom
from difflib import SequenceMatcher


def create_pivot_table(excel_file_path, rename_map):
    try:
        codigo_name = [k for k, v in rename_map.items() if v == "codigo"][0]
        print(codigo_name)
        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False

        wb = excel.Workbooks.Open(excel_file_path)
        ws_data = wb.Sheets(1)
        pivot_sheet = wb.Sheets.Add(After=ws_data)
        pivot_sheet.Name = "PivotTable"

        last_row = ws_data.UsedRange.Rows.Count
        last_col = ws_data.UsedRange.Columns.Count
        data_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(last_row, last_col))
        pivot_table_name = "SummaryPivot"

        pivot_cache = wb.PivotCaches().Create(
            SourceType=win32.constants.xlDatabase,
            SourceData=data_range,
            Version=win32.constants.xlPivotTableVersion15
        )

        pivot_cache.CreatePivotTable(
            TableDestination=pivot_sheet.Cells(1, 1),
            TableName=pivot_table_name,
            DefaultVersion=win32.constants.xlPivotTableVersion15
        )

        pivot_table = pivot_sheet.PivotTables(pivot_table_name)
        headers = [ws_data.Cells(1, i).Value for i in range(2, last_col + 1)]

        for header in headers:
            if header not in [[codigo_name, 'credito', 'debito'], [codigo_name, 'Crédito', 'Débito']]:
                try:
                    pf = pivot_table.PivotFields(header)
                    pf.Orientation = win32.constants.xlRowField
                    pf.Subtotals = [False] * 12
                except:
                    pass
# [['Credito', 'Crédito'], ['Debito', 'Débito']]
        for field in [['Credito', 'Debito'], ['Crédito', 'Débito']]:
            if field in headers:
                try:
                    pf = pivot_table.PivotFields(field)
                    pf.Orientation = win32.constants.xlDataField
                    pf.Function = win32.constants.xlSum
                    pf.Name = field
                except Exception as e:
                    print(f"⚠️ Could not add DataField '{field}': {e}")

        if pivot_table.DataFields.Count == 1:
            try:
                pivot_table.PivotFields("Data").Orientation = win32.constants.xlHidden
            except Exception:
                pass

        if codigo_name in headers:
            try:
                pivot_table.PivotFields(codigo_name).Orientation = win32.constants.xlPageField
            except Exception as e:
                print(f"⚠️ Could not add PageField 'codigo': {e}")

        try:
            pivot_table.RowAxisLayout(win32.constants.xlTabularRow)
            pivot_table.ColumnGrandTotals = False
            pivot_table.RowGrandTotals = False
        except Exception as e:
            print(f"⚠️ Layout customization failed: {e}")

        pivot_sheet.Columns.AutoFit()

        wb.Save()
        wb.Close()
        excel.Quit()
        print("📊 Pivot table generated successfully.")

    except Exception as e:
        print(f"❌ Pivot table creation failed: {e}")




def process_excel_file(df):
    try:
        # === Function: Create Pivot Table ===
      # === Step 2: Token extraction (row-wise, skip 'codigo'-related tokens) ===
        def extract_tokens(desc, codigo):
            tokens = desc.split()
            seen = set()
            return [x for x in tokens if not (x in seen or seen.add(x))]

        df['__pattern_tokens'] = df.apply(lambda row: extract_tokens(row['DESC'], row['codigo']), axis=1)
        df['Pattren'] = df['__pattern_tokens'].apply(lambda tokens: ', '.join(tokens))

        def extract_special_pattern(pattern):
            # Match: start digits (before -), then **, then keep 2/xxxxxx
            match = re.search(r'^(\d+)-[^ ]+\s+TX:\d+\s+(2/\d+)', pattern)
            if match:
                return f"{match.group(1)}**{match.group(2)}"
            return pattern  # keep unchanged if not matched

        def drop_first_pattern(pattern):
            pattern = str(pattern).strip()
            parts = pattern.split()

            removed_info = []
            filtered = []

            for part in parts:
                if part.startswith("/"):  # <-- only keep if it starts with "/"
                    filtered.append(part)
                    continue
                # Rule 1: first 6+ digits followed by 2+ letters (e.g., 837841TT)
                if len(part) >= 8 and part[:6].isdigit() and part[6:].isalpha() and len(part[6:]) >= 2:
                    removed_info.append((part, "Rule: 6+ digits followed by 2+ letters"))
                    continue

                # Rule 2: starts with TX: and only digits after
                if part.startswith("TX:") and part[3:].isdigit():
                    removed_info.append((part, "Rule: TX: + digits"))
                    continue

                # Rule 3: 6+ digits + LR: + digits
                if part[:6].isdigit() and part[6:].startswith("LR:") and part[9:].isdigit():
                    removed_info.append((part, "Rule: 6+ digits + LR: + digits"))
                    continue

                # Rule 4: 6+ digits + LR: + 12+ digits
                if part[:6].isdigit() and part[6:].startswith("LR:") and part[9:].isdigit() and len(part[9:]) >= 12:
                    removed_info.append((part, "Rule: 6+ digits + LR: + 12+ digits"))
                    continue

                # Rule 5: 6+ digits + LR: + alphanumeric
                if part[:6].isdigit() and part[6:].startswith("LR:") and part[9:].isalnum():
                    removed_info.append((part, "Rule: 6+ digits + LR: + alphanumeric"))
                    continue

                # Rule 6: TRJ:**-digit-digit
                if part.startswith("TRJ:**-") and re.match(r"^TRJ:\*\*-\d-\d+$", part):
                    removed_info.append((part, "Rule: TRJ:**-X-X"))
                    continue

                # Rule 7: TRJ:..-digit-digit
                if part.startswith("TRJ:..-") and re.match(r"^TRJ:\.\.-\d-\d+$", part):
                    removed_info.append((part, "Rule: TRJ:..-X-X"))
                    continue

                # Rule 8: 1–2 letters + 2+ digits + 2+ letters + 2+ digits (e.g., S15BUZ612)
                if re.match(r"^[A-Z]{1,2}\d{2,}[A-Z]{2,}\d{2,}$", part):
                    removed_info.append((part, "Rule: 1–2 letters, digits, 2+ letters, 2+ digits"))
                    continue

                # Rule 9: 6+ digits + LR:SPI-PREX + digits
                if part[:6].isdigit() and part[6:].startswith("LR:SPI-PREX") and part[16:].isdigit():
                    removed_info.append((part, "Rule: 6+ digits + LR:SPI-PREX + digits"))
                    continue

                # If no rules matched → keep
                filtered.append(part)

            # Print removed patterns and match source
            for removed_part, rule in removed_info:
                print(f"Removed: '{removed_part}' → {rule}")

            return ' '.join(filtered)


        def fill_pattern_with_referencia(df):
            def mask_last_pattern_if_long_number(df):
                def mask_pattern(pattern_str):
                    # Split into parts
                    patterns = pattern_str.split()
                    if not patterns:
                        return pattern_str
                    
                    # 1️⃣ Mask last part if it contains a long number (10+ digits)
                    last = patterns[-1]
                    digits = re.findall(r'\d+', last)
                    if digits and len(digits[-1]) >= 10:
                        masked = '**' + digits[-1][-4:]
                        patterns[-1] = re.sub(r'\d{7,}', masked, last)

                    # Join back into string
                    result = ', '.join(patterns)

                    # 2️⃣ Replace any special character repeated more than 5 times with "**"
                    result = re.sub(r'([^A-Za-z0-9\s])\1{4,}', '**', result)

                    return result

                df['Pattren'] = df['Pattren'].apply(mask_pattern)

            mask_last_pattern_if_long_number(df)



            def should_replace(pattern):
                pattern = pattern.split()    
                print(f"Checking pattern: '{len(pattern)}'")
                # Empty pattern
        # If pattern has a single token
                if len(pattern) == 1:
                    token = pattern[0]
                    if token.isdigit() and len(token) <= 3:
                        return True
                    if token.isalpha() and len(token) <= 3:
                        return True
                    if ',' not in token and len(token) <= 3:
                        return True
                    if not token.strip():
                        return True
                if len(pattern) < 1:
                    return True
                
            mask_last_pattern_if_long_number(df)    

            for idx, row in df.iterrows():
                pattern = str(row.get('Pattren', '')).strip()
                referencia = str(row.get('Referencia', '')).strip()
                if should_replace(pattern) and referencia:
                    # Split referencia by space and join with comma
                    referencia_patterns = ','.join(referencia.split())
                    df.at[idx, 'Pattren'] = referencia_patterns
                elif pattern:
                    df.at[idx, 'Pattren'] = pattern





        # df['Pattren'] = df['Pattren'].apply(remove_repeated_patterns)
        df['Pattren'] = df['Pattren'].str.replace(',', '', regex=False)

        # === Step 3: Clean numeric fields ===
        for possible_col in [['Credito', 'Crédito'], ['Debito', 'Débito']]:
            # Find which column name actually exists in the dataframe
            col = next((c for c in possible_col if c in df.columns), None)
            if col:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True),
                    errors='coerce'
                ).fillna(0)

        # === Step 4: Save intermediate output to temporary file ===
        df['Pattren'] = df['Pattren'].apply(extract_special_pattern)
        df['Pattren'] = df['Pattren'].apply(drop_first_pattern)
        fill_pattern_with_referencia(df)
        df['Pattren'] = df['Pattren'].str.replace(',', '', regex=False)

        df.drop(columns='__pattern_tokens', inplace=True)
        return df

    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        exit()


def Pre_Processing(df):
    try:
        columns_name = df.columns.tolist()
        print(columns_name)


        required_columns1 = {"Fecha", "Concepto", "Referencia"}
        required_columns2 = {"Fecha valor", "Concepto", "Referencia"}
        required_columns3 = {"Número de documento", "Descripción", "Dependencia"}

#       # Check if the required columns are present in the DataFrame
        if required_columns1.issubset(columns_name):
            rename_map = {
                "Fecha": "codigo",
                "Concepto": "DESC",
                "Referencia": "Referencia"
            }
            df.rename(columns=rename_map, inplace=True)
            df = df.sort_values(by='codigo').reset_index(drop=True)
            df = process_excel_file(df)
            rename_map1 = {
                "codigo": "Fecha",
                "DESC": "Concepto",
                "Referencia": "Referencia"
            }
            df.rename(columns=rename_map1, inplace=True)
            df = remove_empty_columns(df)
            temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            output_path = temp_file.name
            df.to_excel(output_path, index=False)
            temp_file.close()
            print(f"✅ Output saved to temporary file: {output_path}")
            create_pivot_table(output_path, rename_map)            
            return output_path
    

#       # Check if the required columns are present in the DataFrame
        elif required_columns2.issubset(columns_name):
            rename_map = {
                "Fecha valor": "codigo",
                "Concepto": "DESC",
                "Referencia": "Referencia"
            }
            df.rename(columns=rename_map, inplace=True)
            df = df.sort_values(by='codigo').reset_index(drop=True)
            df = process_excel_file(df)
            rename_map1 = {
                "codigo": "Fecha valor",
                "DESC": "Concepto",
                "Referencia": "Referencia"
            }
            df.rename(columns=rename_map1, inplace=True)
            df = remove_empty_columns(df)
            temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            output_path = temp_file.name
            df.to_excel(output_path, index=False)
            temp_file.close()
            print(f"✅ Output saved to temporary file: {output_path}")
            create_pivot_table(output_path, rename_map)            
            return output_path

#       # Check if the required columns are present in the DataFrame
        elif required_columns3.issubset(columns_name):
            rename_map = {
                "Número de documento": "codigo",
                "Descripción": "DESC",
                "Dependencia": "Referencia"
            }
            df.rename(columns=rename_map, inplace=True)
            df = df.sort_values(by='codigo').reset_index(drop=True)
            df = process_excel_file(df)
            rename_map1 = {
                "codigo": "Número de documento",
                "DESC": "Descripción",
                "Referencia": "Dependencia"
            }
            df.rename(columns=rename_map1, inplace=True)
            df = remove_empty_columns(df)
            temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            output_path = temp_file.name
            df.to_excel(output_path, index=False)
            temp_file.close()
            print(f"✅ Output saved to temporary file: {output_path}")
            create_pivot_table(output_path, rename_map)            
            return output_path
        else:
            df = df.sort_values(by='codigo').reset_index(drop=True)
            df = process_excel_file(df)
            rename_map1 = {
                            "codigo": "codigo",
                            "DESC": "DESC",
                            "Referencia": "Referencia"
                        }
            df.rename(columns=rename_map1, inplace=True)
            df = remove_empty_columns(df)
            temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            output_path = temp_file.name
            df.to_excel(output_path, index=False)
            temp_file.close()
            print(f"✅ Output saved to temporary file: {output_path}")
            create_pivot_table(output_path, rename_map)
            return output_path


    except Exception as e:
        print(f"❌ Error loading file: {e}")
        exit()

def remove_empty_columns(df):
    """
    Removes columns from the DataFrame that are entirely empty (all values are NaN or '').
    """
    # Drop columns with all NaN
    df = df.dropna(axis=1, how='all')
    
    # Drop columns with all empty strings after stripping spaces
    df = df.loc[:, ~(df.apply(lambda col: col.astype(str).str.strip() == '').all())]
    
    return df


def process_all_excels_in_folder(input_folder, output_folder):
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Loop through all Excel files in the folder
    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith(('.xlsx', '.xls')):
            input_path = os.path.join(input_folder, file_name)
            print(f"📂 Processing file: {input_path}")

            try:
                df = pd.read_excel(input_path, dtype=str).fillna('')
            except Exception as e:
                print(f"❌ Error reading {file_name}: {e}")
                continue  # Skip this file and move to next

            # Call your Pre_Processing function
            output_path = Pre_Processing(df)  # This returns a temporary output file path

            # Save with same name to output folder
            final_output_path = os.path.join(output_folder, file_name)
            try:
                os.replace(output_path, final_output_path)
                print(f"✅ Final output saved: {final_output_path}")
            except Exception as e:
                print(f"❌ Error saving {final_output_path}: {e}")

            # Remove temporary file if it still exists
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    print(f"🗑️ Temporary file deleted: {output_path}")
                except OSError as e:
                    print(f"❌ Error deleting temporary file {output_path}: {e}")
            else:
                print(f"ℹ️ Temporary file already moved or deleted: {output_path}")



def main():
    # === Step 1: Load Excel file ===
    input_path = r"C:\Users\abhay\OneDrive\Desktop\Data filter\INPUT"
    final_output_dir = r"C:\Users\abhay\OneDrive\Desktop\Data filter\backend\OUTPUT"

    if os.path.isfile(input_path):
        try:
            df = pd.read_excel(input_path, dtype=str).fillna('')
        except FileNotFoundError:
            print(f"❌ Error: The file at {input_path} was not found.")
            exit()

        output_path = Pre_Processing(df)  # This returns the processed file path

        # === Step 6: Save Final Output with Original File Name ===
        original_filename = os.path.basename(input_path)  # e.g., "BROU USD 06 26.xlsx"
        final_output_path = os.path.join(final_output_dir, original_filename)

        os.replace(output_path, final_output_path)
        print(f"📂 Final output file created: {final_output_path}")

        # === Step 7: Cleanup Temp File ===
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
                print(f"🗑️ Temporary file deleted: {output_path}")
            except OSError as e:
                print(f"❌ Error deleting temporary file {output_path}: {e}")
        else:
            print(f"ℹ️ Temporary file already moved or deleted: {output_path}")
    else:
        process_all_excels_in_folder(input_path, final_output_dir)
        
if __name__ == "__main__":
    main()
    