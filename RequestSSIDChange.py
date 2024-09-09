# RequestSSIDChange
# Purpose: Generate Excel sheet for SSID Change requests
# Author: Henry Manning
# Version: 0.0.1

import argparse
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from datetime import datetime

def change_primary_manager(args):
    """Make appropriate updates to the spreadsheet for a primary manager change

    param: args
        Namespace with `change_primary_manager` and `workbook` defined

    return: None 
    """
    # Alias variables
    wb = args.workbook
    new_manager = args.change_primary_manager

    # Modify `Acct Info` sheet
    ws = wb['Acct Info']
    old_manager = ws['B30'].value
    dept = ws['B31'].value
    ws['B28'] = 'Yes'
    ws['B30'] = new_manager

    # Modify `Summary` sheet
    ws = wb['Summary']
    merged_cells_range = ws.merged_cells.ranges
    for merged_cell in merged_cells_range:
        _, top, _, _ = merged_cell.bounds
        if top >= 12:
            merged_cell.shift(0,7)
    ws.insert_rows(12,7)
    cellrange = ws['C12:L17']
    for row in cellrange:
        for cell in row:
            border = copy(cell.border)
            if cell.row == 12:
                border.top = Side(style='medium')
            elif cell.row == 17:
                border.bottom = Side(style='medium')
            if cell.column == 3:
                border.left = Side(style='medium')
            elif cell.column == 12:
                border.right = Side(style='medium')
            cell.border = border
    ws.merge_cells(start_row=12, start_column=3, end_row=17, end_column=12)
    cell = ws['C12']
    cell.alignment = Alignment(vertical='top')

    ws['A12'] = 'Date'
    ws['A13'] = datetime.today().strftime('%m/%d/%Y')
    ws['B12'] = 'REQ'
    ws['C12'] = f'Change primary manager to {new_manager} - previous manager was {old_manager}.'

    # Modify `Previous Ownership` sheet
    ws = wb['Previous Ownership']
    ws.insert_rows(6)
    ws.move_range('A7', rows=-1)
    ws['B6'] = old_manager
    ws['C6'] = dept
    
    return

def parse_args():
    """Define an argparse parser and return the parsed arguments

    return: parsed_args
        Namespace containing parsed arguments
    """
    parser = argparse.ArgumentParser(
                    prog='RequestSSIDChange',
                    description='Generates an Excel spreadsheet for an SSID Change request',
                    epilog='Contact Henry Manning for suggestions. [henry_manning@cinfin.com]')
    
    parser.add_argument('SSID_Name',
                        type=str,
                        help='Name of the SSID being managed')
    
    parser.add_argument('-cpm',
                        '--change-primary-manager',
                        type=str,
                        help='Name of new primary manager')
    
    parser.add_argument('-o',
                        '--output',
                        type=str,
                        default=None)

    parsed_args = parser.parse_args()

    if parsed_args.output is None:
        parsed_args.output = parsed_args.SSID_Name + datetime.today().strftime('%Y-%m-%d') + '.xlsm'

    print(parsed_args.output)

    return parsed_args

def main():
    """Generate new excel sheet
    
    return: status
        0 if success, 1 if any errors are encountered
    """
    status = 0
    args = parse_args()

    try:
        # Generate path to source file
        source_path = args.SSID_Name + '.xlsm'

        # Load workbook
        args.workbook = load_workbook(source_path, read_only=False, keep_vba=True)

        if args.change_primary_manager is not None:
            change_primary_manager(args)

        # Remove broken `legacy_drawing`s
        args.workbook['DB2 UNIX'].legacy_drawing = None
        args.workbook['Mainframe'].legacy_drawing = None
        args.workbook['Other'].legacy_drawing = None

        # Save the file
        args.workbook.save(args.output)

    except Exception as e:
        print(e)
        status = 1

    return status

if __name__ == '__main__':
    main()