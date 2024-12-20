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
$ python3 RequestSSIDChange.py -cpm "Old Manager;New Manager" airflow
    # Modifies spreadsheet for `airflow` SSID to change the primary manager from "Old Manager" to "New Manager"

$ python3 RequestSSIDChange.py -cpm "Old Manager;New Manager" -f ssid_list -o output_dir
    # Modifies spreadsheets for each SSID listed in `ssid_list` file with a new primary manager and places them inside of `output_dir`
```

* **Positional Arguments**
    * `filename`
        * The name of the SSID you want to edit the excel file for, or, if the optional `-f` flag is used, a text file with unique SSIDs on each line
* **Optional Flags**
    * `-cau, --change-authorized-users`
        * Name of previous and new authorized user. Expected format: `<Previous User>;<New User>`
    * `-cm, --change-manager`
        * Name of previous and new manager (either primary or secondary). Expected format: `<Previous Manager>;<New Manager>`
    * `-cpac, --change-primary-account-custodian`
        * Name of previous and new primary account custodian. Expected format: `<Previous Custodian>;<New Custodian>`
    * `-cpm, --change-primary-manager`
        * Name of previous and new primary manager. Expected format: `<Previous Manager>;<New Manager>`
    * `-csm, --change-secondary-manager`
        * Name of previous and new secondary manager. Expected format: `<Previous Manager>;<New Manager>`
    * `-e, --error-logging`
        * Flag to set for error logging
    * `-f, --file-input`
        * Flag to use text file with multiple SSIDs as input, instead of just one
    * `-i, --input-dir`
        * Specify directory to look for SSID spreadsheets in. Defaults to "\\\\wfshq1\\acna\\SSID Forms\\New SSID Forms"
        "\\wfshq1\acna\SSID Forms\New SSID Forms"
    * `-o, --output`
        * Name of file to output to, or directory if using `-f` flag.
    * `-v, --verbose`
        * Flag to set for verbose output