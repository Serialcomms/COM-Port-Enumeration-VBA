# Combo Box Creation

Use one of the two methods below to create and populate a Combo Box  

1.  [VBA subroutine](/combobox/create_vba.md)  

2.  [Create manually](/combobox/create_manual.md)


### Combo Box data list population 
To populate the combo box with COM port names, insert the following code block into
the worksheet, using the combo box name as noted/defined previously.  Example below
uses **ComboBox1** as the sub name prefix and in the **.List** statements. 

From the Ribbon Developer tab, click View Code, right-hand click the worksheet name and select View Code.

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

Check that clicking the combo's drop button causes the com port name(s) to appear in it.   

If applicable, try removing/inserting/enabling any COM Ports and check that the combo box updates correctly. 

The text "NO COM PORTS FOUND" will be displayed in the combo on a PC with no COM Ports available.  

Text defined in declarations section `Public Const TEXT_NO_COM_PORTS As String = "NO COM PORTS FOUND"`
