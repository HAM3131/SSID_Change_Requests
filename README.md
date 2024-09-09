# README
The process for requesting SSID changes is **extremely** cumbersome, and we need to generate hundreds of requests right now. This is an attempt to improve upon that process.

## Requirements
To use this program you must have the `openpyxl` module. Use the following to install it:
```
$ pip install openpyxl
```

## Limitations
"legacy drawings" included on the `DB2 UNIX`, `Mainframe`, and `Other` sheets are broken and must be removed. These "legacy drawings" are just buttons attached to macros which are meant to insert rows and headers automatically in order to simplify/accelerate manual data entry. If using this program, that will no longer be necessary.

## Command line
Example:
```
$ python3 RequestSSIDChange.py -cpm "New Manager" airflow
    # Modifies spreadsheet for `airflow` SSID to change the primary manager to "New Manager"

$ python3 RequestSSIDChange.py -cpm "New Manager" -f ssid_list -o output_dir
    # Modifies spreadsheets for each SSID listed in `ssid_list` file with a new primary manager and places them inside of `output_dir`
```

* **Positional Arguments**
    * `filename`
        * The name of the SSID you want to edit the excel file for, or, if the optional `-f` flag is used, a text file with unique SSIDs on each line
* **Optional Flags**
    * `-cpm, --change-primary-manager`
        * Name of new primary manager
    * `-e, --error-logging`
        * Flag to set for error logging
    * `-f, --file-input`
        * Flag to use text file with multiple SSIDs as input, instead of just one
    * `-i, --input-dir`
        * Specify directory to look for SSID spreadsheets in. Defaults to "\\\\wfshq1\\acna\\SSID Forms\\New SSID Forms"
    * `-o, --output`
        * Name of file to output to, or directory if using `-f` flag.