# pickles

SetUp:
- Open a new excel worksheet
- Go to the view tab and click on the drop down under 'Macros' and click on 'Record Macro'
- Name it anything you want
- For shortcut key input a key you will remember that will run the macros 
  - I use Ctrl+Shift+K for GoThrough()
  - I use Ctrl+Shift+M for DataGather()
- IMPORTANT : Store macro in 'Personal Macro WorkBook'
- Again click on the drop down under 'Macros' and click on 'Stop Recording Macro'
- Go to the developer tab and click on 'Visual Basic' all the way to the left
- Expand 'VBAProject (Personal.XLSB)' then expand 'modules'
- Click on the newly created module and paste DataGather() into the new module and save
- Repeat for GoThrough()

Implementation:
- Have all the files you wish to grab data from and put them in a folder
  - I created a folder named 'FlightData' in "...Machine Learning in Flight Test\Data\MRJ"
- Open Excel
- Use the shortcut for GoThrough() (Ctrl+Shift+K)
- Select the Folder ('FlightData') where the files are stored
-  Use the shortcut again for GoThrough() and selet the same folder again
- Wait... and Click okay when "Done" is displayed
- You should now have a file names 'FlightInfo.csv" in the folder
