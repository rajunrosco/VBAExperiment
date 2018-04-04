Attribute VB_Name = "LocDirectModule"
'LocDirect Constants
Const LocDirectHost As String = "mtl-locdir09.wbiegames.com"
Const LocDirectPort As Long = 50700
Const LocDirectAPI  As String = "api/v1"
Const LocDirectURL As String = "http://" & LocDirectHost & ":" & LocDirectPort & "/" & LocDirectAPI


Dim LocDirectDict As New Scripting.Dictionary
Dim UserName As String
Dim Password As String

Sub SaveCodeModules()

    'This code Exports all VBA modules
    Dim i As Integer, name As String
    
    With ThisWorkbook.VBProject
        For i = .VBComponents.Count To 1 Step -1
            If .VBComponents(i).Type <> vbext_ct_Document Then
                If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                    name = .VBComponents(i).CodeModule.name
                    .VBComponents(i).Export Application.ThisWorkbook.Path & "\\" & name & ".vba"
                End If
            End If
        Next i
    End With

End Sub

Sub LocDirectGetStrings()
    UserName = "SLC_SVC_BUILD"
    Password = "LocDirectSLC"

    Dim AuthBody As String
    Dim response As String
    
    LocDirectDict.RemoveAll
    
    AuthBody = "<?xml version=""1.0"" encoding=""UTF-8""?><EXECUTION client=""API"" version=""1.0""><TASK name=""Login""><OBJECT name=""Security"" /><WHERE><userName>" & UserName & "</userName><password>" & Password & "</password></WHERE></TASK></EXECUTION>"
    
    Dim LocDirectHttpReq As Object
    Set LocDirectHttpReq = CreateObject("Microsoft.XMLHTTP")
    LocDirectHttpReq.Open "POST", LocDirectURL, False
    LocDirectHttpReq.setRequestHeader "Content-Type", "text/xml"
    LocDirectHttpReq.send (AuthBody)
    

    response = LocDirectHttpReq.responseXML.XML
    Dim XDoc As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.LoadXML (response)
    'DEBUG: XDoc.Save ("D:/temp/AuthResponse.xml")
    Set secIds = XDoc.SelectNodes("/EXECUTION/TASK/RESULTSET/DATASETS/DATASET/secId")
    
    Dim secIdString As String
    Dim query As String
    Dim row As Integer
    
    
    Application.StatusBar = "LocDirect GetStrings..."
    
    If (secIds.Length > 0) Then
        secIdString = secIds(0).Text
        query = "<?xml version=""1.0"" encoding=""UTF-8""?><EXECUTION secId=""" & secIdString & """ client=""API"" version=""1.0""><TASK name=""GetStrings""><OBJECT name=""String""><identifierName/><text/></OBJECT><WHERE><projectName>Phoenix</projectName><folderPath>Strings</folderPath><recursive>true</recursive></WHERE></TASK></EXECUTION>"
        LocDirectHttpReq.Open "POST", LocDirectURL, False
        LocDirectHttpReq.send (query)
        
        If LocDirectHttpReq.Status = 200 Then
            'Application.DisplayAlerts = False

            
            XDoc.LoadXML (LocDirectHttpReq.responseXML.XML)
            'DEBUG: XDoc.Save ("D:/temp/GetStringsResponse.xml")
            Set StringList = XDoc.SelectNodes("/EXECUTION/TASK/RESULTSET/DATASETS/Strings/String")
            

            For Each StringNode In StringList
                Set StringIDNode = StringNode.FirstChild
                Set TextNode = StringIDNode.NextSibling
                LocDirectDict.Add StringIDNode.Text, TextNode.Text
            Next StringNode
        End If
    End If
    
    Application.StatusBar = ""
    
    'For Each currentkey In LocDirectDict.Keys
    '    Debug.Print currentkey & " -> " & LocDirectDict(currentkey)
    'Next currentkey
    
End Sub

Sub ProtectLocDirectSheet()
    If Not ActiveSheet.ProtectContents Then
        Columns("E:G").Select
        Selection.Locked = False
        Columns("A:D").Select
        Selection.Locked = True
        ActiveSheet.protect , Contents:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True
    End If

End Sub

Sub UnProtectLocDirectSheet()
    Columns("A:D").Select
    ActiveSheet.unprotect
    Selection.Locked = False
End Sub

Sub RefreshLocDirectData()
    Dim DeleteRowRange As Range
    Dim DiffRange As Range
    Dim ChangedRange As Range
    Dim UpdateRange As Range
    Dim UpdateCell As Range
    Dim SortRange As Range
    Dim WorksheetText As String
    Dim LocDirectText As String
    
    Dim WorksheetStringsDict As New Scripting.Dictionary
    WorksheetStringsDict.RemoveAll
    

    Application.StatusBar = "Refreshing LocDirect Data..."
    
    UnProtectLocDirectSheet
    
    LocDirectGetStrings
    LastRow = Range("B" & Rows.Count).End(xlUp).row
    
    Set DiffRange = Range("A2:A" & LastRow)
    'DiffRange.ClearContents
    For Each DiffCell In DiffRange
        If DiffCell <> "+" Then
            DiffCell.ClearContents
        End If
    Next DiffCell
    
    Set ChangedRange = Range("D2:D" & LastRow)
    ChangedRange.ClearContents
    
    
    Dim DeleteFlag As Boolean
    DeleteFlag = False
    Set DeleteRowRange = Nothing
    
    
    Set UpdateRange = Range("B2:B" & LastRow)
    For Each UpdateCell In UpdateRange
        WorksheetStringsDict.Add UpdateCell.Value, Cells(UpdateCell.row, UpdateCell.Column + 1).Value
        If UpdateCell.Text <> "" Then
            If LocDirectDict.Exists(UpdateCell.Value) Then
                WorksheetText = Cells(UpdateCell.row, UpdateCell.Column + 1).Value
                LocDirectText = LocDirectDict(UpdateCell.Value)
                If WorksheetText <> LocDirectText Then
                    Cells(UpdateCell.row, "A").Value = "*"
                    Cells(UpdateCell.row, "D").Value = LocDirectText & ""
                End If
            Else
                If DeleteRowRange Is Nothing Then
                    Set DeleteRowRange = Rows(UpdateCell.row)
                Else
                    Set DeleteRowRange = Union(DeleteRowRange, Rows(UpdateCell.row))
                End If
                DeleteFlag = True
            End If
        End If
    Next UpdateCell
    
    
    If DeleteFlag Then
        DeleteRowRange.Select
        Selection.EntireRow.Delete
    End If
    
    Dim AddFlag As Boolean
    AddFlag = False
    For Each locdirectkey In LocDirectDict.Keys
        If Not WorksheetStringsDict.Exists(locdirectkey) Then
            AfterLastRow = Range("B" & Rows.Count).End(xlUp).row + 1
            Cells(AfterLastRow, "A").Value = "+"
            Cells(AfterLastRow, "B").Value = locdirectkey
            Cells(AfterLastRow, "C").Value = LocDirectDict(locdirectkey)
            AddFlag = True
        End If
    Next locdirectkey
    
    If AddFlag Then
        LastRow = Range("B" & Rows.Count).End(xlUp).row
        Set SortRange = Range("A2:E" & LastRow)
        SortRange.Sort key1:=Range("B2:B" & LastRow)
    End If
    
    
    
    Application.StatusBar = ""
    MsgBox "LocDirect data refreshed...", vbOKOnly
    
    ProtectLocDirectSheet
    
End Sub

Sub UpdateLocDirect()


    'Debug.Print "UpdateStrings#: " & Len(UpdateStrings)
    'For Each UString In UpdateStrings
    '    Debug.Print UString
    'Next UString
    
    
End Sub


Sub TestSort()


    Dim TestList As Collection
    
    Set TestList = New Collection
    TestList.Add Rows(5377)
    TestList.Add Rows(5374)
    TestList.Add Rows(5372)
    
    
    Dim SortRange As Range
    Dim DeleteRowRange As Range
    

    
    
    For Each deleterow In TestList
        If DeleteRowRange Is Nothing Then
            Set DeleteRowRange = deleterow
        Else
            Set DeleteRowRange = Union(DeleteRowRange, deleterow)
        End If
    Next deleterow
    
    DeleteRowRange.Select
    
    



    
    
    Set SortRange = Range("B5483:C5486")
    SortRange.Sort Range("B5483:B5486")

End Sub
