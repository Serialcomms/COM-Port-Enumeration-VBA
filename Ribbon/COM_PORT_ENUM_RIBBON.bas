Attribute VB_Name = "COM_PORT_ENUM_RIBBON"
'
' https://github.com/Serialcomms/COM-Port-Enumeration-VBA
'
Option Explicit

Public Com_Port_Count As Long
Public Com_Port_Names() As String
Public Com_Port_Numbers() As Long
Public Com_Port_Selected As String
Public Com_Port_Ribbon As IRibbonUI

Public Const TEXT_NO_COM_PORTS As String = "No COM Ports"

Private Declare PtrSafe Function Get_Com_Ports Lib "KernelBase.dll" Alias "GetCommPorts" _
(ByRef Port_Array As Long, ByVal Array_Length As Long, ByRef Port_Count As Long) As Long

Private Const LONG_0 As Long = 0
Private Const LONG_1 As Long = 1
Private Const Max_Port_Count As Long = 255
Private Temp_Port_Numbers(LONG_1 To Max_Port_Count) As Long
'

Public Function Query_Com_Ports() As Long

Dim Port_Number As Long
Dim Port_Ordinal As Long

Get_Com_Ports Temp_Port_Numbers(LONG_1), Max_Port_Count, Com_Port_Count

Erase Com_Port_Names
Erase Com_Port_Numbers
    
If Com_Port_Count > LONG_0 Then
    
    ReDim Com_Port_Names(LONG_1 To Com_Port_Count)
    ReDim Com_Port_Numbers(LONG_1 To Com_Port_Count)
    
    For Port_Ordinal = LONG_1 To Com_Port_Count
    
        Port_Number = Temp_Port_Numbers(Port_Ordinal)

        Com_Port_Numbers(Port_Ordinal) = Port_Number
    
        Com_Port_Names(Port_Ordinal) = "COM" & Port_Number

    Next Port_Ordinal
    
End If

Query_Com_Ports = Com_Port_Count

End Function

Public Function Read_Ribbon_Combo() As String

Application.Volatile

Read_Ribbon_Combo = IIf(Com_Port_Count = LONG_0, vbNullString, Com_Port_Selected)

End Function

Sub InitPortRibbon(ribbon As IRibbonUI)                         'Callback for customUI.onLoad

Set Com_Port_Ribbon = ribbon

Query_Com_Ports

End Sub

Sub PortScan(control As IRibbonControl)                         'Callback for CP_Button onAction

Query_Com_Ports

Com_Port_Ribbon.Invalidate

Com_Port_Selected = vbNullString

Application.Calculate

End Sub

Sub GetButtonLabel(control As IRibbonControl, ByRef ButtonLabel)    'Callback for CP_Button getLabel

Const TEXT_SELECT As String = "Select COM Port"
Const TEXT_DETECT As String = "Detect COM Ports"

ButtonLabel = IIf(Com_Port_Count = LONG_0, TEXT_DETECT, TEXT_SELECT)

End Sub

Sub GetSupertipText(control As IRibbonControl, ByRef SupertipText)   'Callback for CP_Button getSupertipText

Const TEXT_PORTS_AVAILABLE As String = "Com Ports Available = "

Const TEXT_NO_PORTS_FOUND As String = vbCrLf & "No Com ports available " & vbCrLf & vbCrLf & "Click to rescan for new Com ports"

SupertipText = IIf(Com_Port_Count = LONG_0, TEXT_NO_PORTS_FOUND, TEXT_PORTS_AVAILABLE & Com_Port_Count)

End Sub

Sub GetPortText(control As IRibbonControl, ByRef PortText)      'Callback for CP_Selector getPortText

PortText = IIf(Com_Port_Count = LONG_0, TEXT_NO_COM_PORTS, Com_Port_Selected)

Application.Calculate

End Sub

Sub GetPortCount(control As IRibbonControl, ByRef PortCount)      'Callback for CP_Selector getPortCount

PortCount = Query_Com_Ports

Com_Port_Ribbon.Invalidate ' required

End Sub

Sub AddPortID(control As IRibbonControl, index As Integer, ByRef PortID)   'Callback for CP_Selector getPortID

PortID = "Port_ID_" & (index + LONG_1)

End Sub

Sub AddPortLabel(control As IRibbonControl, index As Integer, ByRef PortLabel)    'Callback for CP_Selector getPortLabel

PortLabel = Com_Port_Names(index + LONG_1)

End Sub

Sub GetPortSelection(control As IRibbonControl, PortText As String)          'Callback for CP_Selector onChange

Com_Port_Selected = IIf(Com_Port_Count > LONG_0, PortText, vbNullString)

Application.Calculate

End Sub
