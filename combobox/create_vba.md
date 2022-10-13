
### Create Port Select combo box using supplied VBA

Call supplied VBA sub `create_combo` to create and configure a new ActiveX control combo box. 

VBA creates combo box with name `CP_Selector` as defined in declarations section

`Public Const Port_ComboBox_Name As String = CP_Selector`

Complete the next step to populate the combo box data list with com port names from the host PC, using code block below 
from [Sheet1.bas](Sheet1.bas)

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
