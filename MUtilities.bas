Attribute VB_Name = "MUtilities"
Option Explicit

'---------------------------------------------------------------------------------------
' Method : FolderWithVBAProjectFiles
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Check Folder
'---------------------------------------------------------------------------------------
Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = "C:\Users\Edgar\Downloads" 'WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Method : DeleteVBAModulesAndUserForms
' Author : Edgar de Wit
' Date   : 12-07-18
' Purpose: Delete existing VBA modules
'---------------------------------------------------------------------------------------
Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

Option Explicit
 
 Public Function getFileName(sTitle As String) As String
 
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = sTitle
        .Show
        If .SelectedItems.Count > 0 Then
            getFileName = .SelectedItems(1)
        Else
            getFileName = ""
        End If
    End With

End Function
Function SheetExists(SName As String, Optional ByVal wb As Workbook) As Boolean
'Chip Pearson
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    SheetExists = CBool(Len(wb.Sheets(SName).Name))
End Function
Function FileExists(sPath As String) As Boolean
'Chip Pearson
    Dim FilePath As String
    Dim TestStr As String

    FilePath = sPath
    If FilePath = "" Then
        FileExists = False
        Exit Function
    End If

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function
Function BrowseFolder(Title As String, _
        Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = _
            msoFileDialogViewList) As String
    Dim V As Variant
    Dim InitFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        V = .SelectedItems(1)
        If Err.Number <> 0 Then
            V = vbNullString
        End If
    End With
    BrowseFolder = CStr(V)
End Function
Function GetNameRefersTo(ByRef wkb As Workbook, ByVal TheName As String) As String
    Dim S As String
    Dim HasRef As Boolean
    Dim R As Range
    Dim NM As Name
    Set NM = wkb.Names(TheName)
    On Error Resume Next
    Set R = NM.RefersToRange
    If Err.Number = 0 Then
        HasRef = True
    Else
        HasRef = False
    End If
    If HasRef = True Then
        S = R.Text
    Else
        S = NM.RefersTo
        If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
            ' text constant
            S = Mid(S, 3, Len(S) - 3)
        Else
            ' numeric contant
            S = Mid(S, 2)
        End If
    End If
    GetNameRefersTo = S
End Function


