
#### Create new combo box using supplied VBA

Call the supplied VBA sub `create_combo()` to create and configure a new ActiveX control combo box in worksheet `Sheet1`

To create new combo box in a different worksheet, call the sub with a parameter `create_combo("your_worksheet_name")`

VBA creates a new combo box with name `CP_Selector` as defined in declarations section  

`Public Const Port_ComboBox_Name As String = CP_Selector`

#### Combo Box list population

From the Ribbon Developer tab, click View Code, right-hand click the worksheet name used above and select View Code.

Copy and paste the code block below to populate the combo box data list with com port names from the host PC

```
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

```
