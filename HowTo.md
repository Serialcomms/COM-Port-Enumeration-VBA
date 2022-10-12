# Combo Box Creation

Use one of the methods below to create and use a Combo Box

#### 1. Create combo box using supplied VBA subroutine  

Call sub `create_combo` to create an ActiveX control combo box.  
Combo Box name is defined as `CP_Selector` in declarations section. 


#### 2. Create combo box Manually

Add an ActiveX control combo box to your worksheet from the Excel Ribbon menu  
 * Developer > Insert > ActiveX controls > Combo Box

Right-Hand click the newly-created com to view properties and rename if required. 

Use a combo name of your choice or accept and note the default name given.


### Combo Box population 
To populate the combo box with COM port names, insert the following code block into
the worksheet, using the combo box name as noted/defined previously.  

from Developer tab on Ribbon, click View Code

```
Private Sub ComboBox1_DropButtonClick()

If Query_Com_Ports > 0 Then

    ComboBox1.List = Com_Port_Names()

Else

    ComboBox1.List = Array(TEXT_NO_COM_PORTS)

End If

    Application.Calculate

End Sub
```