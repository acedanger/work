# Useful code

## CSV to XLSX

`PowerShell` script to convert any CSVs found in directory that were created today -

1. saves them as an XLSX
1. renames the first tab
1. moves the XLSX to the parent directory

## Keep Only X Days

Accepts 2 arguments
- `NumDays`
  - integer
  - default is `90`
- `DebugCommand`
  - boolean
  - default is `false`
  
  `PowerShell` script to delete objects in the directory/subdirectory specified that are older than `NumDays`. 
  
