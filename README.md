# Automated .XLSM to .XLS Macro Solution

If you are needing to bulk-convert new Excel .XLSM files to the older .XLS format, this Macro will handle it for you. It automates opening and Saving-As. This does not just "rename" the file, it will specifically handle the "Save-As" function for you automatically.

## Getting Started

You will need Microsoft Excel. Preferrably the latest version to perform this procedure. (At the time of this writing in Oct, 2018)

It's recommended that you SAVE COPIES of your XLSM files into a simple path. The given example is "C:\Test\\". It's best to not attempt running or testing this code in a production folder. Make a backup and run the test from within your backup!

### Editing and Running the Macro

Open your first XLSM spreadsheet (workbook) file in Excel.

Enable macros in Excel. 

Note: (You may need to enable this manually due to default security limiations.) Google for additional guides if necessary. You usually need to enable the "Developer" tab in the Excel ribbon to access macros. 

Open "XLSMtoXLS.vb" from this repo and edit or update the path within the code to point to your actual working path of where you want to run this script. "C:\YourPath\\"

Edit or create a macro in excel, copy the XLSMtoXLS.vb file code into your new macro by overwriting any code in the default macro you created, or by pasting the code into a new macro.

Save the Macro (Ctrl-S) and Run (Play Button) the script. It will convert all XLSM files within the specified directory to XLS format.

I recommend checking the file count within the directory, it should be 2x what it was to begin with. If so, sort by file type and delete the XLSM files manually once the automation process is complete. 

If you want to bulk delete the files auto-magically, you can uncomment out the KILL command within the script. 
