Sub MakeXLS()
Dim sPath As String, sName As String
Dim wkbk As Workbook, sName1 As String
sPath = "C:\Test\"     '<==  change to your path
On Error GoTo Errhandler
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
sName = Dir(sPath & "*.xlsm")
Application.EnableEvents = False
Do While sName <> ""
  sName1 = Replace(sName, ".xlsm", "")
  Application.DisplayAlerts = False
  Set wkbk = Workbooks.Open(sPath & sName)
  wkbk.SaveAs Filename:=sPath & sName1 & ".xls", FileFormat:=xlExcel8
  wkbk.Close SaveChanges:=False
  Application.DisplayAlerts = False
  sName = Dir()
Loop
'<== Kill sPath & "*.xlsm"
Errhandler:
Application.EnableEvents = True
End Sub
