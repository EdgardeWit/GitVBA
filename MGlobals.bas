Attribute VB_Name = "MGlobals"
Option Explicit

'This application needs the following references. Please check before running
'Microsoft Office 16.0 Object Library
'Microsoft Scripting Runtime
'Microsoft Visual Basic for Applications Extensibility 5.3

Public Const xla = "GitVBA"
Public Const web = "https://github.com/EdgardeWit/GitVBA"
Public Const legal = ""
Public Const version = "1.00"

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
