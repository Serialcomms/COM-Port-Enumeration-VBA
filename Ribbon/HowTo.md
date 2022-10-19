## Installing and Testing

####  The main VBA module file should be installed first

<details><summary>VBA Module Installation</summary>
<p>

- Download [COM_PORT_ENUM_RIBBON.bas](COM_PORT_ENUM_RIBBON.bas) to a known location on your PC  
- Open a new Excel document   
- Enter the VBA Environment (Alt-F11)  
- From VBA Environment, view the Project Explorer (Control-R)  
- From Project Explorer, right-hand click and select Import File  
- Import the file COM_PORT_ENUM_RIBBON.bas 
- Check that a new module `COM_PORT_ENUM_RIBBON` is created and visible in the Modules folder
- VBA6 only - delete `PtrSafe` keyword in function definition   
- Close and return to Excel (Alt-Q)  
- IMPORTANT - save document as type Macro-Enabled with a file name of your choice 

  </p>
  </details>
   
#### Excel Ribbon Customisation

<details><summary>Ribbon Customisation</summary>
<p>

- [Ribbon Customisation instructions](Ribbon-HowTo.md)

</p>
</details>


#### Testing

<details><summary>Combo Box and Formula Testing</summary>
<p>
  
Select the COM Port tab and check that a combo box with a label above it are present    

Enter the formula `=Read_Ribbon_Combo()` in any cell to begin

Select a testing scenario below based on the number of COM ports known to be available on the PC.  

<details>
<summary>No COM Ports</summary>
<p>

Check that - 
  
1. Label above combo box is **Detect COM Ports**
2. Hovering over label shows supertip message 'No COM Ports available'
3. Combo box shows message **No COM Ports**
4. Cell with `=Read_Ribbon_Combo()` is blank  
  
</p>
</details>

<details>
<summary>Single COM Port</summary> 
<p>
  
Check that - 
    
1. Label above combo box is **Select COM Port**
2. Hovering over label shows supertip message 'COM Ports available = 1'
3. Single Com Port is available for selection in Combo box
4. Selecting Com port updates cell with selection
5. Clicking **Select COM Port** clears combo box and cell  
  
</p>
</details>

<details>
  
<summary>Multiple COM Ports</summary>
<p>
  
Check that - 
    
1. Label above combo box is **Select COM Port**
2. Hovering over label shows supertip message 'COM Ports available = n'
3. Multiple Com Ports are available for selection in Combo box
4. Selecting a Com port updates cell with selection
5. Selecting a different Com port updates cell with selection 
6. Clicking **Select COM Port** clears combo box and cell   
  
</p>
</details>

<details>  
<summary>Adding / Removing COM Ports</summary>
<p>
  
If possible, check that adding and removing COM Ports updates the worksheet correctly
  
Click the label above the combo box, or re-select the combo box to re-scan available ports.
  
Check that the following items change as expected -   
  
1. Label above Combo Box
2. Label Supertip message
3. Combo box entries
4. Formula cell  
     
##### COM Ports can be removed temporarily by enabling/disabling from the device manager.  
     
</p>
</details>  
  
</p>
</details>
