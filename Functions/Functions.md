#### VBA Functions

| VBA Function                 | Description                                                                                                        |
| ---------------------------- | -------------------------------------------------------------------------------------------------------------------|
| `query_com_ports()`          | Returns number of COM ports, updates Public Variables shown in table below                                         |
| `create_combo()`             | Checks if port selector Combo box exists in Workbook Sheet1 and creates if missing                                 | 
| `create_combo(sheet_name)`   | Checks if port selector Combo box exists in specified sheet and creates if missing                                 |
| `read_combo()`               | Returns COM port selected from combo box in Workbook Sheet1 [^2]                                                   |
| `read_combo(sheet_name)`     | Returns COM port selected from combo box in specified sheet [^2]                                                   |
| `read_ribbon_combo()`        | Returns COM port selected from combo box in customised Excel Ribbon                                                |

#### Public Variables 
| Variable Name              | Variable Type    | Description                                                                                       |
| -------------------------- | -----------------|---------------------------------------------------------------------------------------------------|
| `Com_Port_Count`           | Long             | Number of Com ports returned by `getcommports` [^1]                                               |
| `Com_Port_Names()`         | String Array     | Names of Com ports as text "COM" suffixed by Com Port Number                                      |
| `Com_Port_Numbers()`       | Long Array       | Com port numbers returned by `getcommports` [^1]                                                  |
| `Com_Port_Selected`        | String           | Com port name selected in combo box, or 'no ports found' text                                     |

[^1]: See `getcommports` [documentation](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getcommports) for details of Win32 API function used.
[^2]: Primarily for use within VBA, can also configure combo with `LinkedCell` to update defined worksheet cell directly with port selection.

Notes
1.  Worksheet VBA required in file [`Sheet1.bas`](/Worksheet/Sheet1.bas) to refresh combobox contents list. 
2.  Further development required to use with Excel and other Office applications.
3.  Files in Minimal folder for use with Access and other Office applications.
