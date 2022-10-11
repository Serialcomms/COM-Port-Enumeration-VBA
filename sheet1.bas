Private Sub CP_Selector_Change()

Application.Calculate

End Sub

Private Sub CP_Selector_DropButtonClick()

Query_Com_Ports

Me.CP_Selector.List = Com_Port_Names()

End Sub
