import pandas as pd
import re
import os
import gc
import win32com.client as win32
import tempfile
import pythoncom

def close_existing_excel_instances(file_name):
    """
    Closes any existing Excel workbooks with the specified file_name.
    """
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for wb in excel.Workbooks:
            if wb.Name.lower() == os.path.basename(file_name).lower():
                wb.Close(SaveChanges=False)
        excel.Quit()
        print(f"✅ Closed any existing Excel instances of {os.path.basename(file_name)}")
    except Exception as e:
        print(f"⚠️ Could not close Excel instances: {e}")

def create_pivot_table(excel_file_path):
    """
    Creates a pivot table in a new sheet of the specified Excel file.
    """
    try:
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

        # Create the PivotTable
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

        # Add all fields as RowFields and disable their subtotals
        for header in headers:
            if header not in ['codigo', 'credito', 'debito']:
                try:
                    pf = pivot_table.PivotFields(header)
                    pf.Orientation = win32.constants.xlRowField
                    pf.Subtotals = [False] * 12  # Disable all subtotals
                except:
                    pass

        # Add 'Credito' and 'Debito' as DataFields if present, using their original names
        for field in ['Credito', 'Debito']:
            if field in headers:
                try:
                    pf = pivot_table.PivotFields(field)
                    pf.Orientation = win32.constants.xlDataField
                    pf.Function = win32.constants.xlSum
                    pf.Name = field  # Use original column name, not "Sum of ..."
                except Exception as e:
                    print(f"⚠️ Could not add DataField '{field}': {e}")

        # Hide the default "Data" column if only one data field is present
        if pivot_table.DataFields.Count == 1:
            try:
                pivot_table.PivotFields("Data").Orientation = win32.constants.xlHidden
            except Exception:
                pass

        # Add filter for 'codigo'
        if 'codigo' in headers:
            try:
                pivot_table.PivotFields("codigo").Orientation = win32.constants.xlPageField
            except Exception as e:
                print(f"⚠️ Could not add PageField 'codigo': {e}")

        # Apply tabular layout and remove grand totals for rows/columns
        try:
            pivot_table.RowAxisLayout(win32.constants.xlTabularRow)
            pivot_table.ColumnGrandTotals = False
            pivot_table.RowGrandTotals = False
        except Exception as e:
            print(f"⚠️ Layout customization failed: {e}")

        # Autofit columns for visibility
        pivot_sheet.Columns.AutoFit()

        wb.Save()
        wb.Close()
        excel.Quit()
        print("📊 Pivot table generated, subtotals disabled, tabular layout applied.")
    except Exception as e:
        print(f"❌ An error occurred during pivot table creation: {e}")

# === Step 1: Load and sort Excel file ===
input_path = r"C:\Users\abhay\Downloads\Santander Base de Datos .xlsx"
try:
    df = pd.read_excel(input_path, dtype=str).fillna('')
except FileNotFoundError:
    print(f"❌ Error: The file at {input_path} was not found.")
    exit()

df = df.sort_values(by='codigo').reset_index(drop=True)

# === Step 2: Token extraction from DESC ===
def extract_tokens(desc):
    """
    Extracts specific patterns and tokens from a description string.
    """
    tokens = desc.split()
    extracted = []

    for token in tokens:
        extracted.append(token)
        patterns = re.findall(r'''
            [A-Z]*\d+[A-Z]* |
            TX:\d+               |
            TRJ:[^\s]+           |
            \d{6,}               |
            [A-Z]{2,}\d{2,}      |
            \d+/\d+              |
            -\d+-\d+             |
            \d{15,}              |
            MONTEVIDEO.*?0013
        ''', token, re.VERBOSE | re.IGNORECASE)
        extracted.extend(patterns)

    seen = set()
    return [x for x in extracted if not (x in seen or seen.add(x))]

df['__pattern_tokens'] = df['DESC'].apply(extract_tokens)

# === Step 3: Pattern generation per codigo group ===
def intersect_tokens(group):
    """
    Finds the intersection of tokens within a group.
    """
    all_tokens = group['__pattern_tokens'].tolist()
    if not all_tokens:
        return group.assign(Pattren='')

    common = set(all_tokens[0])
    for token_list in all_tokens[1:]:
        common &= set(token_list)

    ordered_common = [token for token in all_tokens[0] if token in common]
    if not ordered_common:
        ordered_common = all_tokens[0]

    pattern_str = ', '.join(ordered_common)
    return group.assign(Pattren=pattern_str)

# === Step 4: Batch-wise processing ===
batch_size = 100
result_batches = []
for i in range(0, len(df), batch_size):
    batch = df.iloc[i:i+batch_size].copy()
    processed = batch.groupby('codigo', group_keys=False).apply(intersect_tokens)
    result_batches.append(processed)

# === Step 5: Combine all batches ===
final_df = pd.concat(result_batches).reset_index(drop=True)

# === Step 6: Fix duplicate patterns across different codigos ===
pattern_map = final_df.groupby('Pattren')['codigo'].nunique()
duplicate_patterns = pattern_map[pattern_map > 1].index

def prefix_if_duplicate(row):
    """
    Adds a code prefix to patterns that are shared by multiple 'codigo' values.
    """
    if row['Pattren'] in duplicate_patterns:
        return f"*{row['codigo']}*{row['Pattren']}"
    return row['Pattren']

final_df['Pattren'] = final_df.apply(prefix_if_duplicate, axis=1)

# === Step 7: Save final output to a temporary file ===
final_df.drop(columns='__pattern_tokens', inplace=True)

# Clean 'Credito' and 'Debito' columns
for col in ['Credito', 'Debito']:
    final_df[col] = pd.to_numeric(
        final_df[col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True),
        errors='coerce'
    ).fillna(0)

# Create a temporary file to store the processed data
temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
output_path = temp_file.name
final_df.to_excel(output_path, index=False)
temp_file.close()

print(f"✅ Output saved to temporary file: {output_path}")
del final_df
gc.collect()

# === Step 8: Generate Pivot Table ===
# The pivot table creation must happen after the file is saved and closed.
# The `create_pivot_table` function opens, modifies, and saves the file.
create_pivot_table(output_path)
# Save the final output to the specified path
final_output_path = r"C:\Users\abhay\OneDrive\Desktop\Data filter\Final_Token_Patterns_With_CodePrefix.xlsx"
os.replace(output_path, final_output_path)
print(f"📂 Final output file created: {final_output_path}")

# final.drop(columns='__pattern_tokens', inplace=True)  # Removed because 'final' is a string, not a DataFrame
output = os.path.abspath("Final_Token_Patterns_With_CodePrefix.xlsx")



# === Step 9: Cleanup the temporary file ===
try:
    os.remove(output_path)
    print(f"🗑️ Temporary file deleted: {output_path}")
except OSError as e:
    print(f"❌ Error deleting temporary file {output_path}: {e}")