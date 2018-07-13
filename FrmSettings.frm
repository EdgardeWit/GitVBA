VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSettings 
   Caption         =   "Settings"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "FrmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowsePath_Click()
    Dim sPath As String
    sPath = BrowseFolder("Select ApplicationPath")
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    Me.txtApplicationPath = sPath
End Sub

Private Sub cmdCancel_Click()
    mbOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If CheckInput() Then
        mbOK = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Me.txtApplicationPath = gsApplicationPath

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
End Sub
Public Property Get OK() As Boolean
    OK = mbOK
End Property
Public Property Get sPath() As String
    sPath = Me.txtApplicationPath.Value
End Property
Function CheckInput() As Boolean
    Dim bInputOK As Boolean
    
    bInputOK = True
        
    If Not FileExists(Me.txtApplicationPath.Value) Then
        bInputOK = False
        MsgBox "The path " & Me.txtApplicationPath.Value & " can not be found.", vbExclamation
    End If
    
    CheckInput = bInputOK
    
End Function
