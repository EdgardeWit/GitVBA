Attribute VB_Name = "MEntryPoints"
Option Explicit

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



