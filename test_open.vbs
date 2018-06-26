'This script can be called directly from a command prompt to open excel and execute some VBA commands.


Dim ObjExcel, ObjWB

dim document
iRule=msgbox("open")

Set ObjExcel = CreateObject("excel.application")
'vbs opens a file specified by the path below
Set ObjWB = ObjExcel.Workbooks.Open("C:\GitHub\VBAExperiment\test.xlsm")
'either use the Workbook Open event (if macros are enabled), or Application.Run


Set UpdateRange = ObjWB.Worksheets("Phoenix").Range("G3:G6")
For Each UpdateCell In UpdateRange
	ObjExcel.Cells(UpdateCell.row, UpdateCell.column).Value = "BENSON"
Next

ObjWB.Save

iRule=msgbox("close")

ObjWB.Close False
ObjExcel.Quit
Set ObjExcel = Nothing