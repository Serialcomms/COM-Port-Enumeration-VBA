Attribute VB_Name = "COM_PORT_ENUM_MINIMAL"
'
' https://github.com/Serialcomms/COM-Port-Enumeration-VBA
'

Option Explicit

Public Com_Port_Count As Long
Public Com_Port_Names() As String
Public Com_Port_Numbers() As Long

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
