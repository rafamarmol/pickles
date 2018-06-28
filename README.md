# pickles

SetUp:
- Open a new excel worksheet
- Go to the view tab and click on the drop down under 'Macros' and click on 'Record Macro'
- Name it 'CollectData"
- For shortcut key input 'Ctrl+Shift+K'
- IMPORTANT : Store macro in 'Personal Macro WorkBook'
- Again click on the drop down under 'Macros' and click on 'Stop Recording Macro'
- Go to the developer tab and click on 'Visual Basic' all the way to the left
- Expand 'VBAProject (Personal.XLSB)' then expand 'modules'
- Click on the newly created module and paste CollectData() (found in this github) over everything in the module


Implementation:
- Have all the files you wish to grab data from and put them in a folder
  - I created a folder named 'FlightData' in "...Machine Learning in Flight Test\Data\MRJ"
- Open Excel to a untouched blanks workbook
- Use the shortcut for CollectData() (Ctrl+Shift+D)
- Select the Folder ('FlightData') where the files are stored
- It should Open a new Workbook
- Use the shortcut again for CollectData() and selet the same folder again
- Wait... and Follow the on screen instruction until "Done' appears
- You should now have a file names 'FlightInfo.csv" in the folder
