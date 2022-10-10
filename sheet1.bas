Private Sub CP_Selector_Change()

Application.Calculate

End Sub

Private Sub CP_Selector_DropButtonClick()

Me.CP_Selector.List = Get_Port_Names

End Sub
