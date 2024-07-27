Dim ObjExcel, ObjWB

Set ObjExcel = CreateObject("excel.application")
'vbs opens a file specified by the path below


Set NewBook = ObjExcel.Workbooks.Add
' Set to false or remove line to keep excel from showing
objExcel.DisplayAlerts = False
objExcel.Visible = True 
objExcel.Cells(3,3).Value= "Hello from VBS new Script"
objExcel.Cells(10,"M").Value = 6667

NewBook.SaveAs "C:\Temp\testmacro_new.xlsx"
NewBook.Close False

ObjExcel.Quit
Set ObjExcel = Nothing