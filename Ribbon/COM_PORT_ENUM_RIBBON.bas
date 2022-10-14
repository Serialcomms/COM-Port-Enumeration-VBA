Attribute VB_Name = "COM_PORT_ENUM_RIBBON"
'
' https://github.com/Serialcomms/COM-Port-Enumeration-VBA
'
Option Explicit

Public Com_Port_Count As Long
Public Com_Port_Names() As String
Public Com_Port_Numbers() As Long
Public Com_Port_Selected As String

Public Const TEXT_NO_COM_PORTS As String = "NO PORTS"

Private Declare PtrSafe Function Get_Com_Ports Lib "KernelBase.dll" Alias "GetCommPorts" _
(ByRef Port_Array As Long, ByVal Array_Length As Long, ByRef Port_Count As Long) As Long

Private Const LONG_1 As Long = 1
Private Const MAX_Port_Count As Long = 255
Private Temp_Port_Numbers(LONG_1 To MAX_Port_Count) As Long
'

Public Function List_Com_Ports() As Variant

Query_Com_Ports

List_Com_Ports = Com_Port_Names

End Function

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

Public Function Read_Ribbon_Combo() As String

Application.Volatile

Read_Ribbon_Combo = Com_Port_Selected

End Function

Sub InitPortRibbon(ribbon As IRibbonUI) 'Callback for customUI.onLoad

Query_Com_Ports

End Sub

Sub GetPortNames(control As IRibbonControl, ByRef ComboText)                    'Callback for CP_Selector getText

ComboText = IIf(Query_Com_Ports < LONG_1, TEXT_NO_COM_PORTS, Com_Port_Names(LONG_1))

End Sub

Sub GetPortCount(control As IRibbonControl, ByRef ItemCount)                    'Callback for CP_Selector getItemCount

ItemCount = Query_Com_Ports

End Sub

Sub GetPortID(control As IRibbonControl, index As Integer, ByRef ItemID)       'Callback for CP_Selector getItemID

ItemID = "Port_ID_" & (index + LONG_1)

End Sub

Sub GetPortLabel(control As IRibbonControl, index As Integer, ByRef ItemLabel)  'Callback for CP_Selector getItemLabel

ItemLabel = Com_Port_Names(index + LONG_1)

End Sub

Sub GetPortText(control As IRibbonControl, PortName As String)                  'Callback for CP_Selector onChange

Com_Port_Selected = PortName

Application.Calculate

End Sub



