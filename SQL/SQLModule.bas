Attribute VB_Name = "SQLModule"
Public Function SQLTest() As Currency
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strConnString As String
 
    Dim Username As String: Username = Environ("SQLSERVER_USER")
    Dim Password As String: Password = Environ("SQLSERVER_PASS")

    strConnString = Replace(Replace("Provider=SQLOLEDB;ID={Username};PASSWORD={Password};Data Source=MSD-DB-003;" _
                    & "Initial Catalog=MSDUS;Integrated Security=SSPI;", "{Username}", Username), "{Password}", Password)
    Set conn = New ADODB.Connection
    conn.Open strConnString
    
    Dim QueryString As String
    QueryString = "" & _
    "select" & _
    "   c.[Customer_Corporate_Number]," & _
    "   c.[Store Login] " & _
    "from" & _
    "   [MSD-DB-003].[MSDUS].[dbo].[Report_CustomerAddressPivot] as c " & _
    "where" & _
    "   c.[Store Login] <>''" & _
    "   and c.[Customer Balance]<>0" & _
    "   and c.[Customer Number]=[Customer_Corporate_Number]" & _
    "   and c.[Report Region] = 'Domestic'"

    Set rs = conn.Execute(QueryString)
    
    Dim Records As Variant
    Records = rs.GetRows()
    rs.Close

    
    Records = TransposeArray(Records)
    
    Range("A:B").ClearContents
    
    Range(Cells(1, 1), Cells(UBound(Records, 1) + 1, UBound(Records, 2) + 1)).Value = Records

'    If Not IsNumeric(rs.Fields("Customer_Corporate_Number").Value) Then
'        Debug.Print "Empty"
'    Else
'        Debug.Print rs.Fields("Customer_Corporate_Number").Value
'    End If
    
    
End Function


Public Function TransposeArray(myarray As Variant) As Variant
Dim X As Long
Dim Y As Long
Dim Xlower, Xupper As Long
Dim Ylower, Yupper As Long
Dim tempArray As Variant
    Xlower = LBound(myarray, 2)
    Xupper = UBound(myarray, 2)
    Ylower = LBound(myarray, 1)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function
