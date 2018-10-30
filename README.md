Save COPIES of all of your XLSM files into a Path. My example is "C:\Test\"

Be sure to update the path, with your actual path. "C:\YourPath\"

Enable Macros in Excel, Edit/Create a macro and copy the VB file code into your new macro. 

Save and Run the script. It will convert all XLSM files within the specified directory to XLS format.

I recommend checking the file count within the directory, it should be 2x what it was to begin with. If so, sort by file type and delete the XLSM files. 

If you want to bulk delete the files auto-magically, you can uncomment out the KILL command within the script. 
