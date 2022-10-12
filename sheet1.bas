'
' Important = sub name and .List statements should match the combo box name
' sub below is for a combo box with name CP_Selector

Private Sub CP_Selector_DropButtonClick()

If Query_Com_Ports > 0 Then

    CP_Selector.List = Com_Port_Names()

Else

    CP_Selector.List = Array(TEXT_NO_COM_PORTS)

End If

End Sub
