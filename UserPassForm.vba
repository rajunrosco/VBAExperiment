VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserPassForm 
   Caption         =   "LocDirect Credentials"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "UserPassForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserPassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean


Private Declare Function SetEnvironmentVariable Lib "kernel32" _
Alias "SetEnvironmentVariableA" _
(ByVal lpName As String, _
ByVal lpValue As String) As Long



Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub LoginButton_Click()
    Dim user As String
    Dim pass As String
    user = UsernameTextbox.Value
    pass = PasswordTextbox.Value
    
    Dim objUserEnvVars As Object
    Set objUserEnvVars = CreateObject("WScript.Shell").Environment("User")
    objUserEnvVars.Item("LOCDIRECT_USER") = user
    objUserEnvVars.Item("LOCDIRECT_PASSWORD") = pass

    bCancelled = False
    Unload Me
End Sub
