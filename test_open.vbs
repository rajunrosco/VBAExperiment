'This script can be called directly from a command prompt to open excel and execute some VBA commands.

Dim iRule
iRule=msgbox("open")


Dim xl
On Error Resume Next

Do Until status = 429
  Set xl = GetObject(, "excel.application")
  msgbox( TypeName(xl) )
  xl.Close
  xl.Quit
  status = Err.Number
  If status = 0 Then
	
    xl.Quit
  ElseIf status <> 429 Then
    WScript.Echo Err.Number & ": " & Err.Description
    WScript.Quit 1
  End If
Loop


'code to forcibly kill all processes that are excel.exe
' Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
' For Each xl In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'excel.exe'")
 ' msgbox("Terminating id=" & xl.ProcessID )
 ' xl.Terminate
' Next



Dim ObjExcel, ObjWB



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