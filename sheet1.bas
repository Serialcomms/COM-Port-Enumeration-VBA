Private Sub CP_Selector_DropButtonClick()

If Query_Com_Ports > 0 Then

    Me.CP_Selector.List = Com_Port_Names()

Else

    Me.CP_Selector.List = Array(TEXT_NO_COM_PORTS)

End If

End Sub
