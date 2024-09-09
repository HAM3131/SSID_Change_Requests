# README
The process for requesting SSID changes is **extremely** cumbersome, and we need to generate hundreds of requests right now. This is an attempt to improve upon that process.

## Requirements
To use this program you must have the `openpyxl` module. Use the following to install it:
```
$ pip install openpyxl
```

## Limitations
You must also ensure that your spreadsheet is a `.xlsm` file. The `.xls` extension will not work, so you need to save a new version with the correct extension if this applies.

In addition, "legacy drawings" included on the `DB2 UNIX`, `Mainframe`, and `Other` sheets are broken and must be removed. These "legacy drawings" are just buttons attached to macros which are meant to insert rows and headers automatically in order to simplify/accelerate manual data entry. If using this program, that will no longer be necessary.