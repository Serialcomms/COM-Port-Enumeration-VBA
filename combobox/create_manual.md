### Combo Box Creation

#### Create Port Select combo box manually

Add an ActiveX control combo box to your worksheet from the Excel Ribbon menu 

Developer > Insert > ActiveX controls > Combo Box  

Right-Hand click the newly-created combo box to view properties and rename if required.

Use a combo box name of your choice or accept and note the default name given.

Update declarations section `Public Const Port_ComboBox_Name As String = ...` with new combo name.

Complete the next step to populate the combo box data list with com port names from the host PC.
