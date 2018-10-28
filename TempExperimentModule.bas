Attribute VB_Name = "TempExperimentModule"
Public LDRecord As LDRecordType


Public Sub CollectionTest()
    Dim ResultDict As New Scripting.Dictionary
    
    Set LDRecord = New LDRecordType
    
    LDRecord.identifierName_orig = "Key1"
    LDRecord.folderName = "String/Text"
    LDRecord.stringType = "1"
    LDRecord.text = "String1"
    
    ResultDict.Add "key1", LDRecord
    
    Set LDRecord = New LDRecordType
    
    LDRecord.identifierName_orig = "Key2"
    LDRecord.folderName = "String/Text"
    LDRecord.stringType = "2"
    LDRecord.text = "String2"
    
    ResultDict.Add "key2", LDRecord
    
    Dim temp As LDRecordType: Set temp = ResultDict("key2")
    
    Debug.Print temp.identifierName_orig
    ResultDict("key1").text = "String1 Edited"
    
    Debug.Print ResultDict("key1").text
    
    


End Sub



Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Set oExec = oShell.Exec("cmd.exe /c " & sCmd)

    ShellRun = oExec.StdOut.ReadAll

End Function

Function strCompCollection(sourceString As String, ByRef strCollection As Collection) As Boolean
    Dim currentstring As Variant
    For Each currentstring In strCollection
        If currentstring = sourceString Then
            strCompCollection = True
            Exit Function
        End If
    Next currentstring
    strCompCollection = False
End Function

Sub TestInsertPicture()
    ActiveSheet.Shapes.AddPicture _
    Filename:="C:\Users\benso\Pictures\beautiful-bridal-design-291759-main.jpg", _
    linktofile:=msoFalse, savewithdocument:=msoCTrue, _
    Left:=0, Top:=50, Width:=150, Height:=200
    
    
    ActiveSheet.Shapes.AddPicture _
    Filename:="C:\Users\benso\Pictures\beautiful-bridal-design-291759-back.jpg", _
    linktofile:=msoFalse, savewithdocument:=msoCTrue, _
    Left:=100, Top:=50, Width:=150, Height:=200

End Sub

Function TestFindPictures() As Variant()

    Dim picCount As Integer: picCount = 0
    Dim shpTemp As Shape
    Dim shpList As New Collection
    For Each shpTemp In ActiveSheet.Shapes
        Select Case shpTemp.Type
            Case msoLinkedOLEObject
            Case msoLinkedPicture
            Case msoOLEControlObject
            Case msoPicture
                shpList.Add shpTemp.Name
        End Select
    Next shpTemp
    
    If shpList.Count <= 0 Then
        TestFindPictures = Array()
        Exit Function
    End If
    
    
    Dim outputString As String
    Dim temp As Variant
    Dim shpArray() As Variant: ReDim shpArray(0 To shpList.Count - 1)
    
    For Each temp In shpList
        shpArray(picCount) = temp
        outputString = outputString & temp & vbCrLf
        picCount = picCount + 1
    Next temp
    
    'MsgBox outputString, vbOKOnly
    
    TestFindPictures = shpArray
    

End Function

Sub TestDeletePictures()
    Dim shpArray() As Variant
    shpArray = TestFindPictures()
    If (UBound(shpArray) > 0) Then
        MsgBox "Pictures to delete: " & CStr(UBound(shpArray) + 1), vbOKOnly
        Application.ScreenUpdating = False
        ActiveSheet.Shapes.Range(shpArray).Select
        Selection.Delete
        Application.ScreenUpdating = True
    End If
End Sub


Sub Test()

    Dim testcollection As New Collection
    testcollection.Add ("Benson")
    testcollection.Add ("Cleon")
    testcollection.Add ("Lily")
    
    Dim bResult As Boolean
    bResult = strCompCollection("Hello", testcollection)
    bResult = strCompCollection("Cleon", testcollection)
    
    MsgBox ShellRun("ipconfig")

End Sub
