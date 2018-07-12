Dim ObjExcel, ObjWB



Set ObjExcel1 = CreateObject("excel.application")
Set ObjWB = ObjExcel1.Workbooks.Open("C:\GitHub\VBAExperiment\Book1.xlsx")

Set ObjExcel2 = CreateObject("excel.application")
Set ObjWB = ObjExcel2.Workbooks.Open("C:\GitHub\VBAExperiment\Book2.xlsx")

Set ObjExcel3 = CreateObject("excel.application")
Set ObjWB = ObjExcel3.Workbooks.Open("C:\GitHub\VBAExperiment\Book3.xlsx")