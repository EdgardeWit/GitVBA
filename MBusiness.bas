Attribute VB_Name = "MBusiness"
'---------------------------------------------------------------------------------------
' File   : GitVBA
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Import and Export VBA
' Info   : Startpoint https://www.rondebruin.nl/win/s9/win002.htm
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Method : MaintainSettings
' Author : Edgar de Wit
' Date   : 15-07-18
' Purpose: Activate Settings window
'---------------------------------------------------------------------------------------

Sub MaintainSettings()
    Dim frmMaintainSettings As FrmSettings
    Set frmMaintainSettings = New FrmSettings
    With frmMaintainSettings
        .Show
        If .OK Then
            SaveSetting "GitVBA", "Repository", "Path", .sPath
            InitGlobals 'Zorgt er voor dat de nieuwe settings van kracht zijn
        End If
    End With
    Unload frmMaintainSettings
End Sub

'---------------------------------------------------------------------------------------
' Method : CheckOut
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Export every module in active workbook to path
'---------------------------------------------------------------------------------------
Public Sub CheckOut()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent, Message As String

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    Message = MsgBox("Do you want to check out the repository " & gsRepository & "?", vbYesNo, "GitVBA: Check Out")
    If Message = "7" Then Exit Sub
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs gsApplicationPath & ActiveWorkbook.Name, FileFormat:=52
    Application.DisplayAlerts = True

    MsgBox "Check Out is done", vbInformation, "GitVBA: Check out"
End Sub

'---------------------------------------------------------------------------------------
' Method : CheckIn
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Replace all modules in active workbook with repository
'---------------------------------------------------------------------------------------
Public Sub CheckIn()
    Dim wkbTarget As Excel.Workbook, objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File, szTargetWorkbook As String
    Dim szImportPath As String, szFileName As String
    Dim cmpComponents As VBIDE.VBComponents, Message As String

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook. " & _
        "Not possible to import in this workbook", vbCritical, "GitVBA: Check Out failed"
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code", vbCritical, "GitVBA: Check Out failed"
    Exit Sub
    End If
    
    Message = MsgBox("Do you want to check in the repository " & gsRepository & "?", vbYesNo, "GitVBA: Check In")
    
    If Message = "7" Then Exit Sub

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Check In is done", vbInformation, "GitVBA: Check In complete"
End Sub

Sub CreateAddin()
    Dim Message As String
    
    If gsRepository = "" Then InitGlobals
    
    If gsRepository <> "" Then
        Message = MsgBox("Do you want create an add-in from this file to " & gsRepository & "?", vbYesNo, "GitVBA: Create Add-in")
        If Message = "7" Then Exit Sub
        
        
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs gsApplicationPath & GetRightFolder(ActiveWorkbook.FullName), FileFormat:=55
        Application.DisplayAlerts = True
    
        MsgBox "Add-in is created", vbInformation, "GitVBA: Add-in created"
    Else
        Message = MsgBox("Repository path is empty, please check", vbCritical, "GitVBA: No add-in")
        
    End If
    
End Sub
