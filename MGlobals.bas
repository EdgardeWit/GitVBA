Attribute VB_Name = "MGlobals"
Option Explicit

Public gsApplicationPath As String
Public mbOK As Boolean
Public gsRepository As String

'---------------------------------------------------------------------------------------
' Method : InitGlobals
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Setup all globals settings for application
'---------------------------------------------------------------------------------------

Public Sub InitGlobals()
    
    gsApplicationPath = GetSetting("GitVBA", "Repository", "Path")
    If Right$(gsApplicationPath, 1) <> "\" Then gsApplicationPath = gsApplicationPath & "\"
    If Not FileExists(gsApplicationPath) Then
        MsgBox "The database path " & gsApplicationPath & " does not exists. Please change this in settings.", vbExclamation
    End If
    
    gsRepository = GetRightFolder(gsApplicationPath)
    
End Sub
