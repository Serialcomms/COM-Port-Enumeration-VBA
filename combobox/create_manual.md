### Combo Box Creation

#### Create Port Select combo box manually

Add an ActiveX control combo box to your worksheet from the Excel Ribbon menu 

Developer > Insert > ActiveX controls > Combo Box  

Right-Hand click the newly-created combo box to view properties and rename if required.

Use a combo box name of your choice or accept and note the default name given.

Update declarations section `Public Const Port_ComboBox_Name As String = "your_combo_name"` with new combo name.


#### Combo Box list population

From the Ribbon Developer tab, click View Code, right-hand click the worksheet name used above and select View Code.

Insert the following code block into the worksheet code window, using the combo box name as noted/defined previously.  

Example below uses **ComboBox1** as the sub name prefix and in the two **.List** statements. 

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

### Combo box testing

Check that clicking the new combo box's drop button causes com port name(s) or NO COM PORTS FOUND to appear in it.   

If applicable, try removing/inserting/enabling any Com Ports and check that the combo box updates correctly. 

"NO COM PORTS FOUND" will be displayed in the combo on a PC with no Com Ports available.  

Text defined in declarations section `Public Const TEXT_NO_COM_PORTS As String = "NO COM PORTS FOUND"`

