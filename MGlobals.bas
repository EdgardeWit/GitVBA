Attribute VB_Name = "MGlobals"
Option Explicit

Public gsApplicationPath As String
Public mbOK As Boolean


Public Sub InitGlobals()
    
    gsApplicationPath = GetSetting("GitVBA", "Repository", "Path")
    If Right$(gsApplicationPath, 1) <> "\" Then gsApplicationPath = gsApplicationPath & "\"
    If Not FileExists(gsApplicationPath) Then
        MsgBox "The database path " & gsApplicationPath & " does not exists. Please change this in settings.", vbExclamation
    End If
    
End Sub
