Attribute VB_Name = "COM_PORT_ENUMERATION"
'
Option Base 1
Option Explicit

Private Declare PtrSafe Function Get_Com_Ports Lib "KernelBase.dll" Alias "GetCommPorts" (ByRef Port_Array As Long, ByVal Array_Length As Long, ByRef Port_Count As Long) As Long

Public Com_Port_Count As Long
Public Com_Port_Names() As String
Public Com_Port_Selected As String
'
Private Const MAX_PORT_COUNT = 255
Private Temp_Port_Names(1 To MAX_PORT_COUNT) As String
Private Temp_Port_Numbers(1 To MAX_PORT_COUNT) As Long
Private Const Port_Combo_Name As String = "CP_Selector"
'
Public Function Count_Com_Ports() As Long

Get_Com_Ports Temp_Port_Numbers(1), MAX_PORT_COUNT, Com_Port_Count

Count_Com_Ports = Com_Port_Count

End Function

Public Function Get_Port_Names() As Variant

Dim Port_Count As Long
Dim Port_Ordinal As Long

Port_Count = Count_Com_Ports

For Port_Ordinal = 1 To Port_Count

   If Temp_Port_Numbers(Port_Ordinal) > 0 Then

        Temp_Port_Names(Port_Ordinal) = "COM" & CStr(Temp_Port_Numbers(Port_Ordinal))
        
   End If
    
Next Port_Ordinal

Com_Port_Names = Temp_Port_Names

ReDim Preserve Com_Port_Names(Port_Count)

Get_Port_Names = Com_Port_Names

End Function

Public Function Check_Combo(Optional Sheet_Name As String = "Sheet1") As Boolean

Dim WS As Worksheet
Dim CB As Object
Dim CB_Found As Boolean

Set WS = ThisWorkbook.Worksheets(Sheet_Name)

For Each CB In WS.OLEObjects

    If CB.Name = Port_Combo_Name Then
    
        Com_Port_Selected = CB.Object.Value

        Debug.Print CB.Name
        Debug.Print CB.LinkedCell
        Debug.Print CB.Object.Value
        Debug.Print
        
        CB_Found = True
        
        Exit For
        
    End If

Next CB

Check_Combo = CB_Found

End Function

Public Function Read_Combo(Optional Sheet_Name As String = "Sheet1") As String

Application.Volatile

Dim CB As Object

Set CB = ThisWorkbook.Worksheets(Sheet_Name).OLEObjects(Port_Combo_Name)

Com_Port_Selected = CB.Object.Value
        
Read_Combo = Com_Port_Selected

End Function

Public Sub Create_Combo(Optional Sheet_Name As String = "Sheet1")

Dim WS As Worksheet
Dim ComboBox As Object

If Not Check_Combo Then  ' Create new Combo box

    Set WS = ThisWorkbook.Worksheets(Sheet_Name)

    WS.OLEObjects.Add ClassType:="Forms.ComboBox.1"

    Set ComboBox = ThisWorkbook.Worksheets("Sheet1").ComboBox1

    ComboBox.Top = 100
    ComboBox.Left = 80
    ComboBox.Width = 150
    ComboBox.Height = 20
    ComboBox.Name = Port_Combo_Name
    ComboBox.List = Get_Port_Names
    ComboBox.ListIndex = 0
  '  ComboBox.LinkedCell = Com_Port_Selected
    
    
Else

    ' Combo box already exists on this sheet.

End If

End Sub

