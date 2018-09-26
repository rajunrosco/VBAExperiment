Dim ObjExcel, ObjWB

Set ObjExcel = CreateObject("excel.application")
'vbs opens a file specified by the path below
Set ObjWB = ObjExcel.Workbooks.Open("C:\GitHub\VBAExperiment\testmacro.xlsm")
'either use the Workbook Open event (if macros are enabled), or Application.Run

' Set to false or remove line to keep excel from showing
objExcel.Visible = True 

ObjWB.MyMacro


ObjExcel.Cells(3,3).Value= "Hello from VBS"

ObjExcel.Cells(10,"M").Value = 666

'ObjWB.RefreshAll

ObjWB.Save

'dim Msg 
'Msg = MsgBox("done!")

ObjWB.Close False
ObjExcel.Quit
Set ObjExcel = Nothing