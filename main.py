import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import os

# Paths
input_path = os.path.join('Excel Files', 'New Business Scope Sheet - Practice Locations and Providers.xlsx')
output_path = os.path.join('Excel Files', 'Output.xlsx')

def copy_worksheet_values_and_validations(src_ws, tgt_ws):
    # Copy values only (no formatting)
    for row in src_ws.iter_rows():
        for cell in row:
            tgt_ws[cell.coordinate].value = cell.value
    # Copy data validations
    if src_ws.data_validations:
        for dv in src_ws.data_validations.dataValidation:
            tgt_ws.add_data_validation(dv)
    # Copy merged cells
    for merged_range in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged_range))

def main():
    # Load the source workbook
    wb_src = openpyxl.load_workbook(input_path, data_only=False)
    # Create a new workbook for output
    wb_tgt = openpyxl.Workbook()
    # Remove the default sheet if it exists
    if wb_tgt.active and wb_tgt.active.title in wb_tgt.sheetnames:
        wb_tgt.remove(wb_tgt[wb_tgt.active.title])
    # Copy each sheet
    for sheet_name in wb_src.sheetnames:
        src_ws = wb_src[sheet_name]
        tgt_ws = wb_tgt.create_sheet(title=sheet_name)
        copy_worksheet_values_and_validations(src_ws, tgt_ws)
    # Copy named ranges
    for named_range in wb_src.defined_names.values():
        wb_tgt.defined_names.append(named_range)
    # Save the output workbook
    wb_tgt.save(output_path)
    print(f'Copied to {output_path} (values, data validations, named ranges, no formatting)')
    # os.startfile(output_path)  # Moved to end of script

def write_provider_type_substatus_id_formula(output_path, sheet_name='Provider'):
    import openpyxl
    wb = openpyxl.load_workbook(output_path)
    ws = wb[sheet_name]
    max_row = ws.max_row
    col_idx = 61  # 1-based index for 'Provider Type (Substatus) ID'
    assoc_col_idx = 57  # 1-based index for 'Provider Type (Substatus)' (column BE)
    for row in range(2, max_row + 1):
        assoc_value = ws.cell(row=row, column=assoc_col_idx).value
        if assoc_value is not None and str(assoc_value).strip() != '':
            formula = f'=IFERROR(XLOOKUP(BE{row}, ValidationAndReference!Q:Q, ValidationAndReference!P:P, ""), "")'
            ws.cell(row=row, column=col_idx).value = formula
        else:
            ws.cell(row=row, column=col_idx).value = ''
    wb.save(output_path)

def add_board_certification_dropdowns(output_path, sheet_name='Provider'):
    from openpyxl.utils import get_column_letter
    wb = openpyxl.load_workbook(output_path)
    ws = wb[sheet_name]
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    board_cert_cols = []
    for i in range(1, 6):
        col_name = f'Board Certification {i}'
        try:
            col_idx = header.index(col_name) + 1  # openpyxl is 1-indexed
            board_cert_cols.append(col_idx)
        except ValueError:
            print(f"Column '{col_name}' not found in header.")
            continue
    max_row = ws.max_row
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1="=ValidationAndReference!$N$2:$N$299", allow_blank=True)
    dv.error = 'Select a value from the dropdown.'
    dv.errorTitle = 'Invalid Entry'
    dv.prompt = 'Please select a board certification from the list.'
    dv.promptTitle = 'Board Certification'
    ws.add_data_validation(dv)
    for col_idx in board_cert_cols:
        col_letter = get_column_letter(col_idx)
        dv.add(f'{col_letter}2:{col_letter}{max_row}')
    wb.save(output_path)

def add_sub_board_certification_dropdowns(output_path, sheet_name='Provider'):
    from openpyxl.utils import get_column_letter
    wb = openpyxl.load_workbook(output_path)
    ws = wb[sheet_name]
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    sub_board_cert_cols = []
    for i in range(1, 6):
        col_name = f'Sub Board Certification {i}'
        try:
            col_idx = header.index(col_name) + 1  # openpyxl is 1-indexed
            sub_board_cert_cols.append(col_idx)
        except ValueError:
            print(f"Column '{col_name}' not found in header.")
            continue
    max_row = ws.max_row
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1="=ValidationAndReference!$N$2:$N$294", allow_blank=True)
    dv.error = 'Select a value from the dropdown.'
    dv.errorTitle = 'Invalid Entry'
    dv.prompt = 'Please select a sub board certification from the list.'
    dv.promptTitle = 'Sub Board Certification'
    ws.add_data_validation(dv)
    for col_idx in sub_board_cert_cols:
        col_letter = get_column_letter(col_idx)
        dv.add(f'{col_letter}2:{col_letter}{max_row}')
    wb.save(output_path)

def write_board_cert_id_1_formula(output_path, sheet_name='Provider'):
    import openpyxl
    wb = openpyxl.load_workbook(output_path)
    ws = wb[sheet_name]
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    try:
        board_cert_id_col = header.index('Board Cert ID 1') + 1  # 1-based index
        board_cert_col = header.index('Board Certification 1') + 1  # 1-based index (AA)
    except ValueError as e:
        print(f"Column not found: {e}")
        return
    max_row = ws.max_row
    from openpyxl.utils import get_column_letter
    board_cert_col_letter = get_column_letter(board_cert_col)
    for row in range(2, max_row + 1):
        formula = f'=IF(ISBLANK({board_cert_col_letter}{row}),"",INDEX(ValidationAndReference!$AA:$AA,MATCH({board_cert_col_letter}{row},ValidationAndReference!$AB:$AB,0)))'
        ws.cell(row=row, column=board_cert_id_col).value = formula
    wb.save(output_path)

if __name__ == '__main__':
    main()
    # After creating Output.xlsx, copy NPI values
    from NPI import copy_npi_column
    input_sample_path = os.path.join('Excel Files', 'Grow Therapy - Sample Data.xlsx')
    output_path = os.path.join('Excel Files', 'Output.xlsx')
    from Name import copy_name_column
    copy_name_column(input_sample_path, output_path, output_sheet_name='Provider')
    copy_npi_column(input_sample_path, output_path, output_sheet_name='Provider')
    from Gender import copy_gender_column
    copy_gender_column(input_sample_path, output_path, output_sheet_name='Provider')
    from Professional_statement import copy_professional_statement
    copy_professional_statement(input_sample_path, output_path, output_sheet_name='Provider')
    from Langauges import copy_languages
    copy_languages(input_sample_path, output_path, output_sheet_name='Provider')
    from Headshot import copy_headshot_column
    copy_headshot_column(input_sample_path, output_path, output_sheet_name='Provider')
    from PatientsAccepted import copy_patients_accepted
    copy_patients_accepted(input_sample_path, output_path, output_sheet_name='Provider')
    from Professionalsuffix import copy_professional_suffix
    copy_professional_suffix(input_sample_path, output_path, output_sheet_name='Provider')
    # Add hospital affiliation dropdowns
    from Hospitalaff import main as add_hospital_affiliation_dropdowns
    add_hospital_affiliation_dropdowns()
    # Add provider type dropdowns
    from Providertype import add_provider_type_dropdown
    add_provider_type_dropdown()
    # Write formulas to Provider Type (Substatus) ID
    write_provider_type_substatus_id_formula(output_path, sheet_name='Provider')
    # Add Enterprise Scheduling Flag column with dropdown
    from EnterpriseSchedulingFlag import add_enterprise_scheduling_flag_column
    add_enterprise_scheduling_flag_column(output_path, sheet_name='Provider')
    # Add Board Certification 1-5 dropdowns
    add_board_certification_dropdowns(output_path, sheet_name='Provider')
    # Add Sub Board Certification 1-5 dropdowns
    add_sub_board_certification_dropdowns(output_path, sheet_name='Provider')
    # Write Board Cert ID 1 formula
    write_board_cert_id_1_formula(output_path, sheet_name='Provider')
    # Add Location sheet processing
    from Locationsheet import process_location_sheet
    process_location_sheet(input_sample_path, output_path)
    # Open the output file automatically after all processing is complete
    os.startfile(output_path)
    print("The data transposition is complete and is saved as 'Output.xlsx'")
