VERSION 5.00
Begin VB.Form frmSavePath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuList1 
      Caption         =   "List1"
      Visible         =   0   'False
      Begin VB.Menu AddGroup 
         Caption         =   "AddGroup"
      End
      Begin VB.Menu DelGroup 
         Caption         =   "DelGroup"
      End
   End
   Begin VB.Menu mnuList2 
      Caption         =   "List2"
      Visible         =   0   'False
      Begin VB.Menu AddKey 
         Caption         =   "AddKey"
      End
      Begin VB.Menu DelKey 
         Caption         =   "DelKey"
      End
   End
End
Attribute VB_Name = "frmSavePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myIni As clsIniRW

Private Sub AddGroup_Click()
    A$ = InputBox("New Group")
    If A$ = "" Then Exit Sub
    myIni.WriteData A$, "", ""
    ReloadList1
End Sub

Private Sub AddKey_Click()
    A$ = InputBox("New Key")
    If A$ = "" Then Exit Sub
    myIni.WriteData List1.List(List1.ListIndex), A$, ""
    ReloadList2 List1.List(List1.ListIndex)
End Sub

Private Sub DelGroup_Click()
    myIni.WriteData List1.List(List1.ListIndex), vbNullString, vbNullString
    ReloadList1
End Sub

Private Sub DelKey_Click()
    myIni.WriteData List1.List(List1.ListIndex), List2.List(List2.ListIndex), vbNullString
    ReloadList2 List1.List(List1.ListIndex)
End Sub

Private Sub Form_Load()

    Set myIni = frmMain.myIni
    
    ReloadList1
End Sub

Private Sub List1_Click()
    ReloadList2 (List1.List(List1.ListIndex))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuList1
    End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuList2
    End If

End Sub
Private Sub ReloadList1()
    Dim myVariant As Variant
    myVariant = myIni.GetAllSections
    List1.Clear
    For counter = LBound(myVariant) To UBound(myVariant)
        If myVariant(counter) <> "General" Then
            List1.AddItem myVariant(counter)
        End If
    Next

End Sub
Private Sub ReloadList2(SectionName)
    If List1.ListIndex > -1 Then
        Dim myVariant As Variant
        myVariant = myIni.GetAllKeysInSection(CStr(SectionName))
        
        'clear list box
        List2.Clear
        
        'loop thru and add items to list box
        For counter = LBound(myVariant) To UBound(myVariant)
            List2.AddItem myVariant(counter)
        Next
    End If

End Sub
