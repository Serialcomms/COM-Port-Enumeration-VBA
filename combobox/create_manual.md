### Combo Box Creation

#### Create Port Select combo box manually

Add an ActiveX control combo box to your worksheet from the Excel Ribbon menu 

Developer > Insert > ActiveX controls > Combo Box  

Right-Hand click the newly-created combo box to view properties and rename if required.

Use a combo box name of your choice or accept and note the default name given.

Update declarations section `Public Const Port_ComboBox_Name As String = ...` with new combo name.

Complete the next step to populate the combo box data list with com port names from the host PC.

#### Combo Box list population

To populate the combo box with COM port names, insert the following code block into
the worksheet, using the combo box name as noted/defined previously.  Example below
uses **ComboBox1** as the sub name prefix and in the **.List** statements. 

From the Ribbon Developer tab, click View Code, right-hand click the worksheet name used above and select View Code.

```
Private Sub ComboBox1_DropButtonClick()

If Query_Com_Ports > 0 Then

    ComboBox1.List = Com_Port_Names()

Else

    ComboBox1.List = Array(TEXT_NO_COM_PORTS)

End If

    ' Application.Calculate ' optional - if required

End Sub
```
