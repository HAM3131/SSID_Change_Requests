# RequestSSIDChange
# Purpose: Generate Excel sheet for SSID Change requests
# Author: Henry Manning
# Version: 0.0.2

import argparse
import os
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from datetime import datetime

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
            self.source_path = self.filename
            self.tmp_path = os.path.join('tmp', self.filename)
            self.name = name
            self.error_logging = args.error_logging
            self.summary = ''
            self.errored = False

            # Determine output to finally save file
            if args.file_input and args.output is not None:
                if not os.path.exists(args.output):
                    os.makedirs(args.output)
                self.output_path = os.path.join(args.output, name + datetime.today().strftime('%Y-%m-%d') + '.xlsm')
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
            with open(self.source_path, 'rb') as f:
                contents = f.read()
            with open(self.tmp_path, 'wb') as f:
                f.write(contents)

        except Exception as e:
            self.errored = True
            if self.error_logging:
                print(f'SSID.__init__(`{name}`, args): {e}')
        
    def change_primary_manager(self, args):
        """Make appropriate updates to the spreadsheet for a primary manager change

        param: args
            Namespace with `change_primary_manager` and `workbook` defined

        return: None 
        """
        try:
            # Alias variables
            wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
            new_manager = args.change_primary_manager

            # Modify `Acct Info` sheet
            ws = wb['Acct Info']
            old_manager = ws['B30'].value
            dept = ws['B31'].value
            ws['B28'] = 'Yes'
            ws['B30'] = new_manager

            # Modify `Previous Ownership` sheet
            ws = wb['Previous Ownership']
            ws.insert_rows(6)
            ws.move_range('A7', rows=-1)
            ws['B6'] = old_manager
            ws['C6'] = dept
            
            wb.save(self.tmp_path)
            self.summary += f'Change primary manager to {new_manager} - previous manager was {old_manager}. '
        except Exception as e:
            self.errored = True
            if self.error_logging:
                print('SSID.change_primary_manager():', e)
    
    def remove_legacy_drawings(self):
        """Remove the broken legacy drawings on sheets `DB2 UNIX`, `Mainframe`, and `Other`
        """
        # Load the workbook in `tmp` directory
        wb = load_workbook(self.tmp_path, read_only=False, keep_vba=True)
        
        # Remove legacy drawings from broken pages
        wb['DB2 UNIX'].legacy_drawing = None
        wb['Mainframe'].legacy_drawing = None
        wb['Other'].legacy_drawing = None
        
        # Save workbook back to `tmp` folder
        wb.save(self.tmp_path)

    def write_summary(self):
        """Write the summary of all actions taken onto the `Summary` page
        """
        try:
            if self.summary == '':
                raise ValueError('Summary is empty, no changes have been made. Cannot create summary')

            # Modify `Summary` sheet
            ws = load_workbook(self.tmp_path, read_only=False, keep_vba=True)['Summary']
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
        except Exception as e:
            self.errored = True
            if self.error_logging:
                print('SSID.write_summary():', e)
    
    def output(self):
        try:
            if not self.errored:
                with open(self.tmp_path, 'rb') as f:
                    contents = f.read()
                with open(self.output_path, 'wb') as f:
                    f.write(contents)
                print(f'SSID `{self.name}` completed and output at {self.output_path}')
            else:
                print(f'SSID `{self.name}` errored during process - cannot save')
            os.remove(self.tmp_path)
            if len(os.listdir('tmp')) == 0:
                os.rmdir('tmp')
        except Exception as e:
            self.errored = True
            if self.error_logging:
                print('SSID.output():', e)

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
                        help='Name of new primary manager')

    parser.add_argument('-e',
                        '--error-logging',
                        action='store_true',
                        help='Flag to turn on error logging')
    
    parser.add_argument('-f',
                        '--file-input',
                        action='store_true',
                        help='Specify a text file with a different SSID on each line instead of a single SSID to change')
    
    parser.add_argument('-o',
                        '--output',
                        type=str,
                        default=None,
                        help='Filename for output, or directory to output to if using a file as input')

    parsed_args = parser.parse_args()

    return parsed_args

def main():
    """Generate new excel sheet
    
    return: status
        0 if success, 1 if any errors are encountered
    """
    status = 0
    args = parse_args()

    if args.file_input:
        with open(args.filename) as f:
            ssid_names = [line.strip() for line in f.readlines()]
        SSIDs = [SSID(name, args) for name in ssid_names]
    else:
        SSIDs = [SSID(args.filename, args)]
    
    for ssid in SSIDs:
        if args.change_primary_manager is not None:
            ssid.change_primary_manager(args)
        
        ssid.write_summary()
        ssid.output()

    return status

if __name__ == '__main__':
    main()