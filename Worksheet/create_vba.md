
### Combo Box Creation

#### Create a new combo box using supplied VBA

Call the supplied VBA sub `create_combo()` from the VBA Immediate window to create and configure a new ActiveX control combo box in default worksheet `Sheet1`

To create new combo box in a different worksheet, call the sub with a parameter `create_combo("your_worksheet_name")`

VBA creates a new combo box with name `CP_Selector` as defined in declarations section  

`Public Const Port_ComboBox_Name As String = "CP_Selector"`

#### Combo Box dynamic list population

From the Ribbon Developer tab, click View Code, right-hand click the worksheet name used above and select View Code.

Insert the following [code block](/combobox/Sheet1.bas) into the worksheet code window

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
#### Combo Box Testing

Check that clicking the new combo box's drop button causes com port name(s) or NO COM PORTS FOUND to appear in it.   

If applicable, try removing/inserting/enabling any Com Ports and check that the combo box updates correctly. 

"NO COM PORTS FOUND" will be displayed in the combo on a PC with no Com Ports available.  

Text defined in declarations section `Public Const TEXT_NO_COM_PORTS As String = "NO COM PORTS FOUND"`

