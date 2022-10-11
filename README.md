# COM Port Enumeration Excel VBA
## Serial Com Port Enumeration in Excel VBA

#### Functions to assist with end user Com Port selection in Excel

| VBA Function                 | Description                                                                                                        |
| ---------------------------- | -------------------------------------------------------------------------------------------------------------------|
| `query_com_ports()`          | Returns number of COM ports, updates Public variables `Com_Port_Count`, `Com_Port_Names()`, `Com_Port_Numbers()`   |
| `create_combo()`             | Checks if port selector Combo box exists in Workbook Sheet1 and creates if missing                                 | 
| `create_combo(sheet_name)`   | Checks if port selector Combo box exists in specified sheet and creates if missing                                 |
| `read_combo()`               | Returns COM port selected from combo box in Workbook Sheet1                                                        |
| `read_combo(sheet_name)`     | Returns COM port selected from combo box in specified sheet                                                        |


Notes

1.  Worksheet VBA required in file `Sheet1.bas` to refresh combobox contents list. 
2.  Further development required to use with Excel and other Office applications.
3.  See `getcommports` [documentation](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getcommports) for details of Win32 API function used.
