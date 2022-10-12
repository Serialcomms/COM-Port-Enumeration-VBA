### COM Port Enumeration VBA

Function here calls Win32 API to retrieve COM port count and numbers and generate com port names only.   
Further development required to use with Office applications.

#### Public Function

| VBA Function                 | Description                                                                                                        |
| ---------------------------- | -------------------------------------------------------------------------------------------------------------------|
| `query_com_ports()`          | Returns number of COM ports found by Windows API call[^1] , updates Public variables in table below                |


#### Public Variables 

| Variable Name              | Variable Type    | Description                                                                                       |
| -------------------------- | -----------------|---------------------------------------------------------------------------------------------------|
| `Com_Port_Count`           | Long             | Number of Com ports returned by `getcommports` [^1]                                               |
| `Com_Port_Names()`         | String Array     | Names of Com ports as text "COM" suffixed by Com Port Number                                      |
| `Com_Port_Numbers()`       | Long Array       | Com port numbers returned by `getcommports` [^1]                                                  |

[^1]: See `getcommports` [documentation](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getcommports) for details of Win32 API function used.
