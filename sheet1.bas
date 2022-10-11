Private Sub CP_Selector_Change()

Application.Calculate

End Sub


Private Sub CP_Selector_DropButtonClick()

If Query_Com_Ports > 0 Then

    Me.CP_Selector.List = Com_Port_Names()

Else

    Me.CP_Selector.List = Array("NO COM PORTS")

End If

End Sub
