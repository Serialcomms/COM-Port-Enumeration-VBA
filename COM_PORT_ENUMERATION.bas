Attribute VB_Name = "COM_PORT_ENUMERATION"
'
' https://github.com/Serialcomms/COM-Port-Enumeration-VBA
'
Option Explicit

Public Com_Port_Count As Long
Public Com_Port_Names() As String
Public Com_Port_Numbers() As Long
Public Com_Port_Selected As String
Public Const Port_ComboBox_Name As String = "CP_Selector"
Public Const TEXT_NO_COM_PORTS As String = "NO COM PORTS FOUND"

Private Declare PtrSafe Function Get_Com_Ports Lib "KernelBase.dll" Alias "GetCommPorts" _
(ByRef Port_Array As Long, ByVal Array_Length As Long, ByRef Port_Count As Long) As Long

Private Const LONG_1 As Long = 1
Private Const MAX_Port_Count As Long = 255
Private Temp_Port_Numbers(LONG_1 To MAX_Port_Count) As Long

Public Function Query_Com_Ports() As Long

Dim Port_Ordinal As Long

Get_Com_Ports Temp_Port_Numbers(LONG_1), MAX_Port_Count, Com_Port_Count

If Com_Port_Count < LONG_1 Then

    Erase Com_Port_Names
    Erase Com_Port_Numbers

Else

    ReDim Com_Port_Names(LONG_1 To Com_Port_Count)
    ReDim Com_Port_Numbers(LONG_1 To Com_Port_Count)

    For Port_Ordinal = LONG_1 To Com_Port_Count

        Com_Port_Numbers(Port_Ordinal) = Temp_Port_Numbers(Port_Ordinal)
    
        Com_Port_Names(Port_Ordinal) = "COM" & Temp_Port_Numbers(Port_Ordinal)

    Next Port_Ordinal

End If

Query_Com_Ports = Com_Port_Count

End Function

Public Function Read_Combo(Optional Sheet_Name As String = "Sheet1") As String

Application.Volatile

Dim CB As Object

Set CB = ThisWorkbook.Worksheets(Sheet_Name).OLEObjects(Port_ComboBox_Name)

Com_Port_Selected = CB.Object.Value
        
Read_Combo = Com_Port_Selected

End Function

Public Function Check_Combo(Optional Sheet_Name As String = "Sheet1") As Boolean

Dim WS As Worksheet
Dim CB As Object
Dim CB_Found As Boolean

Set WS = ThisWorkbook.Worksheets(Sheet_Name)

For Each CB In WS.OLEObjects

    If CB.Name = Port_ComboBox_Name Then
    
        Com_Port_Selected = CB.Object.Value
        
        CB_Found = True
        
        Exit For
        
    End If

Next CB

Check_Combo = CB_Found

End Function

Public Sub Create_Combo(Optional Sheet_Name As String = "Sheet1")

Dim WS As Worksheet
Dim ComboBox As Object

    If Not Check_Combo(Sheet_Name) Then  ' Create new Combo box

    Set WS = ThisWorkbook.Worksheets(Sheet_Name)

    WS.OLEObjects.Add ClassType:="Forms.ComboBox.1"

    Set ComboBox = ThisWorkbook.Worksheets(Sheet_Name).ComboBox1

    ComboBox.Top = 100
    ComboBox.Left = 80
    ComboBox.Width = 150
    ComboBox.Height = 20
    ComboBox.Name = Port_ComboBox_Name
    ComboBox.List = IIf(Query_Com_Ports > 0, Com_Port_Names, Array(TEXT_NO_COM_PORTS))
    ComboBox.ListIndex = 0
  '  ComboBox.LinkedCell = "A5"
    
    
Else

    ' Combo box already exists on this sheet.

End If

End Sub

