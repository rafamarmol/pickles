# pickles

SetUp:
- Open a new excel worksheet
- Go to the view tab and click on the drop down under 'Macros' and click on 'Record Macro'
- Name it 'DataSeperator"
- For shortcut key input 'Ctrl+Shift+Z'
- IMPORTANT : Store macro in 'Personal Macro WorkBook'
- Again click on the drop down under 'Macros' and click on 'Stop Recording Macro'
- Go to the developer tab and click on 'Visual Basic' all the way to the left
- Expand 'VBAProject (Personal.XLSB)' then expand 'modules'
- Click on the newly created module and paste DataSeperator() (found in this github) over everything in the module


Implementation:
- Open the file you wish to extract data from
- Use the shortcut for DataSeperator() (Ctrl+Shift+Z)
- Wait... 
- You should now have a 3 files named "FTI" , "BusData", "ASix" 
