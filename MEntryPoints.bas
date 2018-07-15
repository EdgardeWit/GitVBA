Attribute VB_Name = "MEntryPoints"
'---------------------------------------------------------------------------------------
' File   : MEntryPoints
' Author : Edgar de Wit
' Date   : 15-07-18
' Purpose: Start points of this application
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Method : ribbonClick
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Catch up module for the toolbar
'---------------------------------------------------------------------------------------

Public Sub ribbonClick(control As IRibbonControl)

On Error GoTo ErrorHandler

    Select Case control.Tag
        Case "ImportModules"
            ImportModules
        Case "ExportModules"
            ExportModules
        Case "Settings"
            MaintainSettings
    End Select

ErrorExit:
     Exit Sub
    
ErrorHandler:
    MsgBox "An unexpected error occurred, the program was terminated" & _
        Chr(13) & "The full error description is: " & _
        Err.Description, vbCritical, "GitVBA"
    Resume ErrorExit

End Sub



