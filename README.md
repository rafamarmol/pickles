# pickles

SetUp:
- Open a new excel worksheet
- Go to the view tab and click on the drop down under 'Macros' and click on 'Record Macro'
- Name it 'CollectDataBETA'
- For shortcut key input 'Ctrl+Shift+Z'
- IMPORTANT : Store macro in 'Personal Macro WorkBook'
- Again click on the drop down under 'Macros' and click on 'Stop Recording Macro'
- Go to the developer tab and click on 'Visual Basic' all the way to the left
- Expand 'VBAProject (Personal.XLSB)' then expand 'modules'
- Click on the newly created module and paste over everything in the module CollectDataBETA() (found in this github) 


Implementation:
- move all the flights yopu wish to extract data from to a folder of your chosing
- Use the shortcut for CollectDataBETA() (Ctrl+Shift+Z)
- Select the file where all the flights are located
- Follow on screen instructions (If clipboard storage notification pops up, you can click no)
- Wait... 
- You should now have a a file called "FlightInfo.csv"
#DALE
