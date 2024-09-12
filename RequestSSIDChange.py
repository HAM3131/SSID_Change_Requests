# RequestSSIDChange
# Purpose: Generate Excel sheet for SSID Change requests
# Author: Henry Manning
# Version: 0.0.5

import argparse
import os
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from datetime import datetime
import win32com.client as win32
import time

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.DisplayAlerts=False

class SSID:
    def __init__(self, name, args):
        """Initialize SSID object

        param: name
            The name of the SSID to be modified
        param: args
            Namespace containing command line arguments
        """
        try:
            self.filename = name + '.xlsm'
            self.tmp_path = os.path.join('tmp', self.filename)
            self.name = name
            self.error_logging = args.error_logging
            self.logs = ''
            self.summary = ''
            self.errored = False

            # Determine source path
            self.source_path = os.path.join(args.input_dir, self.filename)
            # Excel file with name of SSID isn't present in input directory. Check for dir
            if not os.path.isfile(self.source_path):
                ssid_folder = os.path.join(args.input_dir, name)
                max_mtime = 0
                if os.path.isdir(ssid_folder) and len(os.listdir(ssid_folder)) > 0:
                    for entry in os.listdir(os.path.join(args.input_dir, name)):
                        full_path = os.path.join(ssid_folder, entry)
                        mtime = os.stat(full_path).st_mtime
                        if mtime > max_mtime:
                            max_mtime = mtime
                            self.source_path = full_path
                else:
                    raise ValueError(f'No path found to source file for SSID `{name}`')
            
            # Determine output to finally save file
            if args.file_input and args.output is not None:
                if not os.path.exists(args.output):
                    os.makedirs(args.output)
                self.output_path = os.path.join(args.output, name + '_' + datetime.today().strftime('%Y-%m-%d') + '.xlsm')
            elif args.output is not None:
                self.output_path = args.output
                if not self.output_path.endswith('.xlsm'):
                    self.output_path += '.xlsm'
            else:
                self.output_path = name + datetime.today().strftime('%Y-%m-%d') + '.xlsm'

            # Create `tmp` directory if not present 
            if not os.path.isdir('tmp'):
                os.makedirs('tmp')

            # Copy file into tmp folder for working
            if not self.source_path.endswith('.xlsm'):
                copy_excel_as_xlsm(self.source_path, os.path.join(os.getcwd(), self.tmp_path))
            else:
                with open(self.source_path, 'rb') as f:
                    contents = f.read()
                with open(self.tmp_path, 'wb') as f:
                    f.write(contents)

        except ValueError as e:
            self.errored = True
            error = f'ERROR: SSID.__init__(`{name}`, args): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)
        
        else:
            self.logs += f'SSID `{self.name}` initialized successfully\n'
        
    def change_primary_manager(self, args):
        """Make appropriate updates to the spreadsheet for a primary manager change

        param: args
            string with previous and new manager separated by semicolon: `<Previous Manager>;<New Manager>`

        return: None 
        """
        try:
            # Alias variables
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
            old_manager, new_manager = args.split(';')

            # Modify `Acct Info` sheet
            ws = wb['Acct Info']
            if not old_manager.lower() == ws['B30'].value.lower():
                raise ValueError(f'previous primary manager = `{ws["B30"].value}`, expected `{old_manager}`')
            dept = ws['B31'].value
            ws['B28'] = 'Yes'
            ws['B30'] = new_manager

            # Modify `Previous Ownership` sheet
            self.modify_previous_ownership(wb, 'Primary Manager', name=old_manager, dept=dept)
            
            wb.save(self.tmp_path)
            self.summary += f'Change primary manager to {new_manager} - previous manager was {old_manager}. '
        
        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.change_primary_manager(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)
        
        else:
            self.logs += f'Primary manager changed from `{old_manager}` to `{new_manager}` for SSID `{self.name}`\n'
    
    def change_secondary_manager(self, args):
        """Make appropriate updates to the spreadsheet for a secondary manager change

        param: args
            string with previous and new managers separated by semicolon: `<Previous Manager>;<New Manager>`

        return: None 
        """
        try:
            # Alias variables
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
            old_manager, new_manager = args.split(';')

            # Modify `Acct Info` sheet
            ws = wb['Acct Info']
            if not old_manager.lower() == ws['B32'].value.lower():
                raise ValueError(f'previous secondary manager = `{ws["B32"].value}`, expected `{old_manager}`')
            dept = ws['B33'].value
            ws['B28'] = 'Yes'
            ws['B32'] = new_manager

            # Modify `Previous Ownership` sheet
            self.modify_previous_ownership(wb, 'Secondary Manager', name=old_manager, dept=dept)
            
            wb.save(self.tmp_path)
            self.summary += f'Change secondary manager to {new_manager} - previous manager was {old_manager}. '
        
        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.change_secondary_manager(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)
            return False
        
        else:
            self.logs += f'Secondary manager changed from `{old_manager}` to `{new_manager}` for SSID `{self.name}`\n'
            return True

    def change_manager(self, args):
        """Make appropriate updates to the spreadsheet for a primary manager change

        param: args
            Namespace with `change_manager` defined

        return: None 
        """
        try:
            # Alias variables
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
            old_manager, _ = args.change_manager.split(';')

            ws = wb['Acct Info']
            primary_manager = ws['B30'].value
            if primary_manager is None:
                primary_manager = ''
            secondary_manager = ws['B32'].value
            if secondary_manager is None:
                secondary_manager = ''
            if primary_manager.lower() == old_manager.lower():
                self.logs += f'change_manager() selected `primary manager` for SSID `{self.name}`\n'
                self.change_primary_manager(args.change_manager)
            elif secondary_manager.lower() == old_manager.lower():
                self.logs += f'change_manager() selected `secondary manager` for SSID `{self.name}`\n'
                self.change_secondary_manager(args.change_manager)
            else:
                raise ValueError(f'Neither primary nor secondary manager matches expected previous manager: `{old_manager}`')
        
        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.change_manager(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)

    def remove_legacy_drawings(self):
        """Remove the broken legacy drawings on sheets `DB2 UNIX`, `Mainframe`, and `Other`
        """
        try:
            # Load the workbook in `tmp` directory
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
            
            broken_drawings = ['DB2 UNIX', 'DB2 AIX', 'Mainframe', 'Other']
            # Remove legacy drawings from broken pages
            for sheet in list(set(broken_drawings) & set([sheet.title for sheet in wb._sheets])):
                wb[sheet].legacy_drawing = None
            
            # Save workbook back to `tmp` folder
            wb.save(self.tmp_path)

        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.remove_legacy_drawings(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)
        
        else:
            self.logs += f'Legacy drawings removed from SSID `{self.name}` successfully\n'

    def write_summary(self):
        """Write the summary of all actions taken onto the `Summary` page
        """
        try:
            if self.summary == '':
                raise ValueError('Summary is empty, no changes have been made. Cannot create summary')

            # Modify `Summary` sheet
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
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
            ws['C12'] = self.summary

            # Save workbook
            wb.save(self.tmp_path)

        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.write_summary(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)

        else:
            self.logs += f'Summary written for SSID `{self.name}` successfully\n'

    def modify_previous_ownership(self, wb, field, name='', dept=''):
        if not 'Previous Ownership' in wb:
            wb.create_sheet('Previous Ownership')
            ws = wb['Previous Ownership']
            ws['A4'] = 'TITLE'
            ws['B4'] = 'NAME'
            ws['C4'] = 'DEPARTMENT'
            ws['D4'] = 'SIGNATURE & DATE'
            ws['A6'] = 'Primary Manager'
            ws['A13'] = 'Secondary Manager'
            ws['A19'] = 'Primary Account Custodian'
            ws['A26'] = 'Secondary Account Custodian'
            ws['A30'] = 'Authorized User(s)'
            ws['A43'] = 'Authorized User\'s Manager'
            ws['A47'] = 'ISO Representative(s)'
        
        ws = wb['Previous Ownership']
        row_num = None
        for cell in ws['A']:
            if cell.value == field:
                row_num = int(cell.row)
                break
        
        if row_num is None:
            raise ValueError(f'ERROR: `{field}` field not found in COL(A) of `Previous Ownership` sheet')
                    
        ws.insert_rows(row_num)
        ws.move_range(f'A{row_num}', rows=-1)
        ws[f'B{row_num}'] = name
        ws[f'C{row_num}'] = dept


    def output(self):
        """Save the spreadsheet to the final output destination
        """
        try:
            if not self.errored:
                with open(self.tmp_path, 'rb') as f:
                    contents = f.read()
                with open(self.output_path, 'wb') as f:
                    f.write(contents)
                print(f'SSID `{self.name}` completed and output at {self.output_path}')
                os.remove(self.tmp_path)
            else:
                print(f'SSID `{self.name}` errored during process - cannot save')
            if os.path.isdir('tmp') and len(os.listdir('tmp')) == 0:
                os.rmdir('tmp')
        except ValueError as e:
            self.errored = True
            error = f'ERROR: `{self.name}` - SSID.output(): {e}'
            self.logs += error + '\n'
            if self.error_logging:
                print(error)
        
        else:
            self.logs += f'SSID `{self.name}` output successfully\n'
        
    def log(self):
        """Write logs for SSID
        """
        path = 'logs'
        if self.errored:
            path = os.path.join(path, 'fail')
        else:
            path = os.path.join(path, 'success')
        
        if not os.path.isdir(path):
            os.makedirs(path)

        path = os.path.join(path, self.name + '.log')

        with open(path, 'w') as f:
            f.write(self.logs)

def copy_excel_as_xlsm(source_path, output_path):
     try:
        wb = excel.Workbooks.Open(source_path)
        wb.SaveAs(output_path, 52)
        wb.Close(SaveChanges=False)
     except IOError:
        print('Error')

def parse_args():
    """Define an argparse parser and return the parsed arguments

    return: parsed_args
        Namespace containing parsed arguments
    """
    parser = argparse.ArgumentParser(
                    prog='RequestSSIDChange',
                    description='Generates an Excel spreadsheet for an SSID Change request',
                    epilog='Contact Henry Manning for suggestions. [henry_manning@cinfin.com]')
    
    parser.add_argument('filename',
                        type=str,
                        help='Name of the SSID being managed')
                        
    parser.add_argument('-cpm',
                        '--change-primary-manager',
                        type=str,
                        help='Name of previous and new primary manager. Expected format: `<Previous Manager>;<New Manager>`')
    
    parser.add_argument('-csm',
                        '--change-secondary-manager',
                        type=str,
                        help='Name of previous and new secondary manager. Expected format: `<Previous Manager>;<New Manager>`')
    
    parser.add_argument('-cm',
                        '--change-manager',
                        type=str,
                        help='Name of previous and new manager (either primary or secondary). Expected format: `<Previous Manager>;<New Manager>`')

    parser.add_argument('-e',
                        '--error-logging',
                        action='store_true',
                        help='Flag to turn on error logging')
    
    parser.add_argument('-f',
                        '--file-input',
                        action='store_true',
                        help='Specify a text file with a different SSID on each line instead of a single SSID to change')

    parser.add_argument('-i',
                        '--input-dir',
                        type=str,
                        default='\\\\wfshq1\\acna\\SSID Forms\\New SSID Forms',
                        help='Path to dir which SSID spreadsheets are in. Defaults to `\\\\wfshq1\\acna\\SSID Forms\\New SSID Forms`')
    
    parser.add_argument('-o',
                        '--output',
                        type=str,
                        default=None,
                        help='Filename for output, or directory to output to if using a file as input')

    parsed_args = parser.parse_args()

    # Ensure `-cpm, -csm, -cm` flags are not used together
    if len([flag for flag in [parsed_args.change_primary_manager, parsed_args.change_secondary_manager, parsed_args.change_manager] if flag is not None]) > 1:
        raise ValueError(f'INPUT ERROR: Multiple flags for changing managers used. Use just one of `-cpm, --change-primary-manager`, `-csm, --change-secondary-manager`, and `-cm, --change-manager`')

    # Ensure `-cpm, --change-primary-manager` value has proper syntax, including a semicolon separating new and old managers
    if parsed_args.change_primary_manager is not None and not ';' in parsed_args.change_primary_manager:
        raise ValueError(f'`INPUT ERROR: -cpm, --change-primary-manager` flag used without proper syntax: `{parsed_args.change_primary_manager}`\n'
                        +f'\tExpected: `<Previous Manager>;<New Manager>`')
    
    # Ensure `-csm, --change-secondary-manager` value has proper syntax, including a semicolon separating new and old managers
    if parsed_args.change_secondary_manager is not None and not ';' in parsed_args.change_secondary_manager:
        raise ValueError(f'`INPUT ERROR: -csm, --change-secondary-manager` flag used without proper syntax: `{parsed_args.change_secondary_manager}`\n'
                        +f'\tExpected: `<Previous Manager>;<New Manager>`')
    
    # Ensure `-cm, --change-manager` value has proper syntax, including a semicolon separating new and old managers
    if parsed_args.change_manager is not None and not ';' in parsed_args.change_manager:
        raise ValueError(f'INPUT ERROR: `-cm, --change-manager` flag used without proper syntax: `{parsed_args.change_manager}`\n'
                        +f'\tExpected: `<Previous Manager>;<New Manager>`')

    return parsed_args

def get_ssid_list(args):
    """Generate list of SSID objects from input

    param: args
        Namespace including `file_input` and `filename`
    
    return: SSIDs
        list of SSID objects
    """
    if args.file_input:
        with open(args.filename) as f:
            ssid_names = [line.strip() for line in f.readlines()]
        SSIDs = [SSID(name, args) for name in ssid_names]
    else:
        SSIDs = [SSID(args.filename, args)]
    
    return SSIDs

def execute_changes(args):
    """Initialize list of SSIDs, make appropriate edits to spreadsheets, and save

    param: args
        Namespace including `change_primary_manager` and other key details of command line arguments
    """
    print('\033[1;34m***Initializing SSID objects***\033[22;0m')
    SSIDs = get_ssid_list(args)

    if args.change_primary_manager is not None:
        print('\033[1;34m***Changing primary managers***\033[22;0m')
        [ssid.change_primary_manager(args.change_primary_manager) for ssid in SSIDs if not ssid.errored]

    if args.change_secondary_manager is not None:
        print('\033[1;34m***Changing secondary managers***\033[22;0m')
        [ssid.change_secondary_manager(args.change_primary_manager) for ssid in SSIDs if not ssid.errored]

    if args.change_manager is not None:
        print('\033[1;34m***Changing primary manager***\033[22;0m')
        [ssid.change_manager(args) for ssid in SSIDs if not ssid.errored]
    
    print('\033[1;34m***Writing change summaries***\033[22;0m')
    [ssid.write_summary() for ssid in SSIDs if not ssid.errored]

    # Remove legacy drawings
    [ssid.remove_legacy_drawings() for ssid in SSIDs if not ssid.errored]
    
    print('\033[1;34m***Saving successfully modified files***\033[22;0m')
    [ssid.output() for ssid in SSIDs if not ssid.errored]

    [ssid.log() for ssid in SSIDs]

    successful_edits = len([ssid for ssid in SSIDs if not ssid.errored])
    print(f'\033[1;32mFile editing completed -\033[22;0m {successful_edits}/{len(SSIDs)} edited successfully')

def main():
    """Generate new excel sheet
    
    return: status
        0 if success, 1 if any errors are encountered
    """
    status = 0

    try:
        args = parse_args()
        execute_changes(args)
    except ValueError as e:
        print(e)

    return status

if __name__ == '__main__':
    main()

excel.Quit()