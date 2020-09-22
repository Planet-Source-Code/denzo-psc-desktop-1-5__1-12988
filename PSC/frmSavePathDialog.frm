VERSION 5.00
Begin VB.Form frmSavePathDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SavePathDialog"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Skip"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   255
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Library Name"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Library Path"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmSavePathDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myIni As New clsIniRW
Dim LibPath() As String
Public Skip As Boolean

Private Sub Combo1_Click()
    On Error Resume Next
    Text1.Text = LibPath(Combo1.ListIndex)
    Text2.Text = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Command1_Click()
    a$ = Text2.Text
    If a$ = "" Then Exit Sub
    On Error Resume Next
        MkDir Text1.Text
        If Err > 0 And Err <> 75 Then
            MsgBox Error
        Else
            If Right$(Text1.Text, 1) <> "\" Then
                Text1.Text = Text1.Text + "\"
            End If
            myIni.WriteData "Libraries", a$, Text1.Text
        End If
    On Error GoTo 0
    Form_Load
    Combo1.Text = a$
    For i = 0 To Combo1.ListCount - 1
        If Combo1.Text = Combo1.List(i) Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub Command2_Click()
    Dim BrowseFolder As New CBrowseFolder
    
    With BrowseFolder
        .hwnd = Me.hwnd
        .Prompt = "Select the default destination folder:"
        .ShowStatus = True
        .Folder = Text1.Text
        .IncludeFiles = False
        .BrowseForFSDirs = True
        .BrowseForFSAncestor = False
        .BrowseForPrinter = False
        .StartUpPosition = vbStartUpOwner
        .Show
        SavePath.Text = .Folder
    End With
    

End Sub

Private Sub Command3_Click()
    myIni.DeleteData "Libraries", Combo1.Text, LibPath(Combo1.ListIndex)
    Form_Load
End Sub

Private Sub Command4_Click()
    Skip = False
    Me.Hide
End Sub

Private Sub Command5_Click()
    Skip = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    myIni.File_Name = App.Path + "\DenzoPSC.ini"
    Text1.Text = myIni.GetValue("General", "SavePath")
    Combo1.Clear
    Combo1.AddItem "Root"
    Combo1.Text = "Root"
    Dim a As Variant
    a = myIni.GetAllKeysInSection("Libraries")
    If UBound(a) > -1 Then
        For i = LBound(a) To UBound(a)
            Combo1.AddItem a(i)
        Next i
    End If
    Dim b As Variant
    b = myIni.GetAllKeysValuesInSection("Libraries")
    ReDim LibPath(0)
    ReDim LibPath(200)
    For j = 0 To Combo1.ListCount - 1
        If "Root" = Combo1.List(j) Then
            Exit For
        End If
    Next j
    LibPath(j) = myIni.GetValue("General", "SavePath")
    If UBound(b) > -1 Then
        For i = LBound(b) To UBound(b)
            'ReDim Preserve LibPath(1 + i)
            d$ = myIni.GetValue("Libraries", CStr(a(i)))
            For j = 0 To Combo1.ListCount - 1
                If CStr(a(i)) = Combo1.List(j) Then
                    Exit For
                End If
            Next j
            LibPath(j) = d$
            'SetFolderIcon d$, App.Path + "\PSC.ico", 0
        Next i
    End If
    On Error Resume Next
    a = myIni.GetAllKeysInSection("Libraries")
    Err = 5
    MkDir Text1.Text + "Libraries"
    Do While Err <> 0 And UBound(a) > 0
        Err = 0
        For i = LBound(a) To UBound(a)
            MkDir LibPath(i)
            If Err = 75 Then Err = 0
        Next i
        
    Loop
    On Error GoTo 0
End Sub

