VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   Caption         =   "PSC Desktop"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2528
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1005
      ButtonWidth     =   1429
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Download"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Browser"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Libraries"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   10560
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   14
      Top             =   7560
      Width           =   4575
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   315
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdLoadTree 
      Caption         =   "cmdLoadTree"
      Height          =   735
      Left            =   10200
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmeMain 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   11668
         _Version        =   393217
         Indentation     =   365
         LineStyle       =   1
         Style           =   6
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   9855
         ExtentX         =   17383
         ExtentY         =   11245
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin InetCtlsObjects.Inet intMain 
      Left            =   240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTxtBox 
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":2842
   End
   Begin RichTextLib.RichTextBox txtURL 
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":28C4
   End
   Begin VB.Frame frmeNav 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   6520
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox SavePath 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "C:\"
         Top             =   240
         Width           =   6375
      End
      Begin VB.CommandButton Save 
         Caption         =   "Save Checked"
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdPrevpage 
      Caption         =   "Prev Page"
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next Page"
      Height          =   495
      Left            =   9360
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   10185
      Picture         =   "frmMain.frx":2946
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   10200
      Picture         =   "frmMain.frx":2C50
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   5415
   End
   Begin VB.Menu nmuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuBrowser 
         Caption         =   "&Browser"
      End
   End
   Begin VB.Menu mnuCategories 
      Caption         =   "&Categories"
      Begin VB.Menu mnuNewest 
         Caption         =   "Newest"
      End
      Begin VB.Menu mnuBest 
         Caption         =   "Best"
      End
      Begin VB.Menu mnuNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "CodingStandards"
         Index           =   0
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Complete Applications"
         Index           =   1
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Custom Controls/ Forms/ Menus"
         Index           =   2
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Data Structures"
         Index           =   3
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Databases/ Data Access/ DAO/ ADO"
         Index           =   4
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "DDE"
         Index           =   5
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Debugging and Error Handling"
         Index           =   6
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "DirectX"
         Index           =   7
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Files/ File Controls/ Input/ Output"
         Index           =   8
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Games"
         Index           =   9
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Graphics"
         Index           =   10
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Internet/ HTML"
         Index           =   11
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Jokes/ Humor"
         Index           =   12
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Math/ Dates"
         Index           =   13
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Microsoft Office Apps/VBA"
         Index           =   14
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Miscellaneous"
         Index           =   15
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "OLE/ COM/ DCOM/ Active-X"
         Index           =   16
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Registry"
         Index           =   17
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Sound/MP3"
         Index           =   18
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "String Manipulation"
         Index           =   19
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "VB function enhancement"
         Index           =   20
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Windows API Call/ Explanation"
         Index           =   21
      End
      Begin VB.Menu mnuCategoriesSpec 
         Caption         =   "Windows System Services"
         Index           =   22
      End
   End
   Begin VB.Menu Search 
      Caption         =   "&Search"
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu Libraries 
         Caption         =   "Libraries..."
      End
      Begin VB.Menu UseSavingKey 
         Caption         =   "UseSavingKey"
      End
      Begin VB.Menu EditSavingKey 
         Caption         =   "EditSavingKey"
      End
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Calls API to change the look and feel of the command buttons
Option Compare Text
Dim Root As Node
Dim PageNumber As String
Dim PathUrl$
Dim QueryType$
Dim StopMe As Boolean
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Dim WithEvents Tray As TrayIcon
Attribute Tray.VB_VarHelpID = -1
Dim Progress As New ProgressBar
Dim NotLoaded As Boolean
Public myIni As New clsIniRW
Dim Downloading As Boolean
'Dim DownloadCompleted As Boolean


Private Sub cmdExtract_Click()
  Dim aString As String
  Static Executing As Boolean
    If Executing Then
        MsgBox "Still executing command!"
        Exit Sub
    End If
    Executing = True
    frmMain.MousePointer = 11
    lblStatus.Caption = "Status: Loading Souce from " & txtURL.Text
    'Opens URL and outputs the souce into the richtextbox
    On Error Resume Next
    'aString = intMain.OpenURL(txtURL.Text)
    DownloadCompleted = False
    WebBrowser1.Navigate txtURL.Text
    'Do While DownloadCompleted = False
    '    DoEvents
    'Loop
    Dim u As Long
    'WebBrowser1.Document.documentElement.innerHTML = ""
    Do While u < 600000 'WebBrowser1.LocationURL <> txtURL.Text
        u = u + 1
        DoEvents
    Loop
    'Do While Not WebBrowser1.ReadyState >= READYSTATE_COMPLETE And u < 6000000
    'aString = WebBrowser1.Document.documentElement.innerHTML
    u = 0
    Do While ((InStr(WebBrowser1.Document.documentElement.innerHTML, "<!--interstitial-->") = 0) And (u < 6000000)) 'Or (Len(WebBrowser1.Document.documentElement.innerHTML) < 47000)  'And (WebBrowser1.ReadyState > READYSTATE_LOADED)
        'aString = WebBrowser1.Document.documentElement.innerHTML
        DoEvents
        u = u + 1
        If Len(WebBrowser1.Document.documentElement.innerHTML) = 0 Then
            If u > 300000 Then
                Err.Raise 200000, "PSCDescktop", "Unable to connect to server!"
                Exit Do
            End If
        End If
    Loop
    errString = Err.Description
    Err = 0
    aString = WebBrowser1.Document.documentElement.innerHTML
    If Err > 0 Or u >= 6000000 Then
        MsgBox "Unable to connect to server!" + " Try later!"
        StopMe = True
    End If
    If InStr(aString, "error '80040e14'") <> 0 Then
        WebBrowser1.Navigate txtURL.Text
        Do While WebBrowser1.Busy
            DoEvents
        Loop
    End If
    On Error GoTo 0
    RichTxtBox.Text = aString
    If RichTxtBox.Text <> "" Then
        lblStatus.Caption = "Status: Completed loading source from " & txtURL.Text
    Else
        lblStatus.Caption = "Status: Problem occured with " & txtURL.Text & " either you mistyped the url or the website doesn't exsist."
    End If
    Executing = False
    frmMain.MousePointer = 0
End Sub



Private Sub cmdFolder_Click()
    Dim BrowseFolder As New CBrowseFolder
    
    With BrowseFolder
        .hwnd = Me.hwnd
        .Prompt = "Select the default destination folder:"
        .ShowStatus = True
        .Folder = SavePath.Text
        .IncludeFiles = False
        .BrowseForFSDirs = True
        .BrowseForFSAncestor = False
        .BrowseForPrinter = False
        .StartUpPosition = vbStartUpOwner
        .Show
        SavePath.Text = .Folder
    End With
    
    Set BrowseFolder = Nothing
End Sub

Private Sub cmdLoadTree_Click()
    Dim Loaded(500) As String
    On Error Resume Next
    
    Open App.Path + "\Downloaded.dat" For Input As #2
        If Err = 0 Then
            i = 0
            Tot = 0
            Do While Not EOF(2)
                Input #2, Loaded(i)
                i = i + 1
                If i > 500 Then
                    Tot = 500
                    i = i - 500
                End If
                If Err > 0 Then i = i - 1: Exit Do
            Loop
        End If
    Close #2
    On Error GoTo 0
    
    If Tot = 500 Then
        TotLoaded = 500
    Else
        TotLoaded = i - 1
    End If
    Dim nNode As Node
    Dim tmpNode As Node
    TreeView1.Nodes.Clear
    Set Root = TreeView1.Nodes.Add(, , "Root", QueryType$)
    Root.BackColor = &H80000016
    Root.ForeColor = &H8000000E
    Root.Bold = True
    a$ = RichTxtBox.Text
    'Open "C:\tmp.txt" For Output As #1
    '    Print #1, A$
    'Close #1
    StartElem = InStr(a$, "<!--descrip-->")
    EndElem = InStr(StartElem + 1, a$, "<!--descrip-->")
    If EndElem = 0 Then
        NotLoaded = True
        Exit Sub
    End If
    NotLoaded = False
    'b$ = ClearSpecialChar(Mid$(A$, StartElem, EndElem - StartElem))
    Do While StartElem <> 0
        b$ = ClearSpecialChar(Mid$(a$, StartElem, EndElem - StartElem))
        Codenumber = Val(Mid$(b$, InstrEnd(b$, "txtCodeId.")))
        'CodeName = Trim(GetInsideString(B$, "alt=" + Chr$(34), Chr$(34)))
        'CodeName = Trim(GetInsideString(b$, "alt=", "src"))
        CodeName = Trim(GetInsideString(b$, "alt=", "src"))
        CodeName = Replace(CodeName, "<", " ")
        CodeName = Replace(CodeName, ">", " ")
        CodeName = Replace(CodeName, "\", " ")
        CodeName = Replace(CodeName, "/", " ")
        CodeName = Replace(CodeName, ":", " ")
        CodeName = Replace(CodeName, "*", " ")
        CodeName = Replace(CodeName, "?", " ")
        CodeName = Replace(CodeName, ",", " ")
        CodeName = Replace(CodeName, "&amp;", " ")
        CodeName = Replace(CodeName, ";", " ")
        CodeName = Replace(CodeName, "&", " ")
        'CodeName = Replace(CodeName, "(", " ")
        CodeName = MakeProper(CodeName)
        CodeName = Replace(CodeName, " ", "")
        'CodeName = Replace(CodeName, "-", "-")
        CodeName = Replace(CodeName, Chr$(34), "")
        On Error GoTo err1
        ShowCodeUrl = "/xq" + Trim(GetInsideString(b$, "/xq", "ShowCode.htm")) + "ShowCode.htm"
        b$ = Mid$(b$, InStr(b$, "<!--code compat-->"))
        CodeCompat = Trim(GetInsideString(b$, "<!--code compat-->", "</TD>"))
        b$ = Mid$(b$, InStr(b$, "<!--level-->"))
        CodeLevel = Trim(GetInsideString(b$, "<!--level-->", " /"))
        b$ = Mid$(b$, InStr(b$, "/"))
        CodeAuth = Trim(GetInsideString(b$, "<BR>", "<"))
        If CodeAuth = "" Then CodeAuth = Trim(GetInsideString(b$, Chr$(34) + ">", "<"))
        b$ = Mid$(b$, InStr(b$, "><!--views/date submitted-->"))
        CodeDate = Trim(GetInsideString(b$, "since<BR>", "</TD>"))
        b$ = Mid$(b$, InStr(b$, "<!description>"))
        CodeDescription = Trim(GetInsideString(b$, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "<HR>"))
        On Error Resume Next
        If InStr(CodeDescription, "href") > 0 Then
            CodeDescription = Mid$(CodeDescription, 1, InStr(CodeDescription, "<a") - 1)
        End If
        If TreeView1.Nodes("ID" + Format(Codenumber)).Text <> "" Then
            TreeView1.Nodes.Remove ("ID" + Format(Codenumber))
        End If
        On Error GoTo 0
        Set nNode = TreeView1.Nodes.Add(Root, tvwChild, "ID" + Format(Codenumber), CodeName)
        nNode.Bold = True
        For i = 0 To TotLoaded
            If Codenumber = Loaded(i) Then
                nNode.ForeColor = &H80000013
                Exit For
            End If
        Next i

        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "DESCR" + Format(Codenumber), CodeDescription)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "URL" + Format(Codenumber), "http://www.planet-source-code.com" + ShowCodeUrl)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "COMPAT" + Format(Codenumber), CodeCompat)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "LEVEL" + Format(Codenumber), CodeLevel)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "AUTH" + Format(Codenumber), CodeAuth)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "DATE" + Format(Codenumber), CodeDate)
        Set tmpNode = TreeView1.Nodes.Add(nNode, tvwChild, "NUMBER" + Format(Codenumber), Codenumber)
err1:
        StartElem = InStr(EndElem, a$, "<!--descrip-->")
        EndElem = InStr(StartElem + 1, a$, "<!--descrip-->")
        If EndElem = 0 Then EndElem = InStr(StartElem + 1, a$, "</TABLE>")
    Loop
    Root.Expanded = True

End Sub

Private Sub cmdNextPage_Click()
    p = Val(PageNumber) + 1
    PageNumber = Format(p)
    txtURL.Text = PathUrl$
    cmdExtract_Click
    cmdLoadTree_Click

End Sub

Private Sub cmdPrevpage_Click()
    p = Val(PageNumber) - 1
    If p < 1 Then p = 1
    PageNumber = Format(p)
    txtURL.Text = PathUrl$
    cmdExtract_Click
    cmdLoadTree_Click

End Sub

Private Sub EditSavingKey_Click()
    frmSavePath.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Static Message As Long
    'Message = X / Screen.TwipsPerPixelX
    'Select Case Message
    'Case WM_RBUTTONUP:
    '    PopupMenu TrayMenu
    'End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'MsgBox Format(UnloadMode)
    If UnloadMode = 0 Then
        Cancel = True
        WindowState = 1
    End If
End Sub

Private Sub Form_Resize()
    'Exit Sub
    On Error Resume Next
    If WindowState = 1 Then
        ResizeForm Me
        Me.Hide
        'showTrayIcon Form, "PSC Desktop"
        Tray.Add
    Else
        'hideTrayIcon Form
        Me.Show
        Tray.Remove
        ResizeForm Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'hideTrayIcon Form
    Tray.Remove
    SaveSetting "PSC Desktop", "User Defaults", "Save Path", SavePath.Text
End Sub

Private Sub Libraries_Click()
    frmSavePathDialog.Show 1
End Sub

Private Sub mnuAbout_Click()
    MsgBox "This is a little piece of code written to help www.Planet-Source-Code.com  to become more and more useful. Thanks everybody for their work. Bye Denzo!"
End Sub

Private Sub mnuBest_Click()
    PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories.-1/txtMaxNumberOfEntriesPerPage.10/optSort./chkThoroughSearch./blnTopCode.True/blnNewestCode.False/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.10/intMaxNumberOfEntriesPerPage.10/intLastRecordInRecordset.8986/chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria./cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.10/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
    QueryType$ = "Best"
    txtURL.Text = PathUrl$
    cmdExtract_Click
    cmdLoadTree_Click
End Sub

Private Sub mnuBrowser_Click()
    'WebBrowser1.Visible = Not WebBrowser1.Visible
    frmMain.TreeView1.Visible = Not frmMain.TreeView1.Visible
    mnuBrowser.Checked = Not mnuBrowser.Checked
End Sub

Private Sub mnuCategoriesSpec_Click(Index As Integer)
    PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories.%Category%/txtMaxNumberOfEntriesPerPage.100/optSort.DateDescending/chkThoroughSearch./blnTopCode.False/blnNewestCode.False/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.10/intMaxNumberOfEntriesPerPage.10/intLastRecordInRecordset.128/chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria./cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.10/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
    QueryType$ = mnuCategoriesSpec(Index).Caption
    
    Select Case Index
        Case 0
            Category = "43"
        Case 1
            Category = "27"
        Case 2
            Category = "4"
        Case 3
            Category = "33"
        Case 4
            Category = "6"
        Case 5
            Category = "28"
        Case 6
            Category = "26"
        Case 7
            Category = "44"
        Case 8
            Category = "3"
        Case 9
            Category = "38"
        Case 10
            Category = "46"
        Case 11
            Category = "34"
        Case 12
            Category = "40"
        Case 13
            Category = "37"
        Case 14
            Category = "42"
        Case 15
            Category = "1"
        Case 16
            Category = "29"
        Case 17
            Category = "36"
        Case 18
            Category = "45"
        Case 19
            Category = "5"
        Case 20
            Category = "25"
        Case 21
            Category = "39"
        Case 22
            Category = "35"
        Case Else
            Category = "43"
    End Select

    PageNumber = "1"
    PathUrl$ = Replace(PathUrl$, "%Category%", Category)
    txtURL.Text = PathUrl$
    cmdExtract_Click
    cmdLoadTree_Click

End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuNewest_Click()
    PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories.-1/txtMaxNumberOfEntriesPerPage.50/optSort.DateDescending/chkThoroughSearch./blnTopCode.False/blnNewestCode.True/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.50/intMaxNumberOfEntriesPerPage.50/intLastRecordInRecordset./chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria./cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.50/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
    'PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories.-1/txtMaxNumberOfEntriesPerPage.50/optSort.DateDescending/chkThoroughSearch./blnTopCode.False/blnNewestCode.True/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.50/intMaxNumberOfEntriesPerPage.50/intLastRecordInRecordset.10640/chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria./cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.50/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
    PathUrl$ = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=50&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=1"
    QueryType$ = "Newest"
    txtURL.Text = PathUrl$
    cmdExtract_Click
    cmdLoadTree_Click
    PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories.-1/txtMaxNumberOfEntriesPerPage.50/optSort.DateDescending/chkThoroughSearch./blnTopCode.False/blnNewestCode.True/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.50/intMaxNumberOfEntriesPerPage.50/intLastRecordInRecordset./chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria./cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.50/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
End Sub

Private Sub mnuRefresh_Click()
    cmdExtract_Click
    cmdLoadTree_Click

End Sub

Private Sub Quit_Click()
    Unload Me
End Sub

Private Sub Save_Click()
    Dim nNode As Node
    Dim ID As String
    Downloading = True
    Icon = Image1(0).Picture

    Progress.Clear
    If Right$(SavePath.Text, 1) <> "\" Then SavePath.Text = SavePath.Text + "\"
    For Each nNode In TreeView1.Nodes
        If nNode.Checked And Mid$(nNode.Key, 1, 2) = "ID" Then
            NumNodi = NumNodi + 1
        End If
    Next nNode
    i = 0
    StopMe = False
    For Each nNode In TreeView1.Nodes
        If StopMe Then Exit For
        If nNode.Checked And Mid$(nNode.Key, 1, 2) = "ID" Then
            ID = Mid$(nNode.Key, 3)
            txtURL.Text = TreeView1.Nodes("URL" + ID).Text
            cmdExtract_Click
            Progress.Percent = ((i + 0.5) / (NumNodi)) * 100
            a$ = RichTxtBox.Text
            Load frmSavePathDialog
            frmSavePathDialog.Caption = nNode.Text
            frmSavePathDialog.Label3 = TreeView1.Nodes("DESCR" + ID).Text
            'nNode.Text = Mid$(nNode.Text, 1, 15)
            
            frmSavePathDialog.Show 1
            If Not frmSavePathDialog.Skip Then
                On Error Resume Next
                FilePath = frmSavePathDialog.Text1.Text + nNode.Text
                MkDir FilePath
                If Err = 75 Then
                    u = 1
                    Do While Err = 75
                        Err = 0
                        FilePath = frmSavePathDialog.Text1.Text + nNode.Text + Format(u)
                        u = u + 1
                        MkDir FilePath
                    Loop
                Else
                    If Err > 0 Then
                        MsgBox Error
                    End If
                End If
                Err = 0
                Open FilePath + "\" + nNode.Text + ".htm" For Output As #1
                    If Err > 0 Then
                        MsgBox Err.Description
                        Close #1
                        GoTo NextItem
                    End If
                    Print #1, a$
                Close #1
                On Error Resume Next
                'ZipName = Trim(GetInsideString(A$, "http://www.planet-source-code.com/upload/ftp/CODE_UPLOAD", ".zip"))
                'ZipName = Trim(GetInsideString(Mid$(a$, InStr(a$, "title end")), "<a href=" + Chr$(34) + "/upload/ftp/", "."))
                zipname = Trim(GetInsideString(Mid$(a$, InStr(a$, "<!zip file>")), "<a href=" + Chr$(34) + "/vb/scripts/ShowZip", Chr$(34)))
                zipname = Replace(zipname, "&amp;", "&")
                If Err = 0 Then
                    'GetZip "http://www.planet-source-code.com/upload/ftp/" + ZipName + ".zip", frmSavePathDialog.Text1.Text + nNode.Text + Format(u) + "\" + nNode.Text + ".zip"
                    GetZip "http://www.planet-source-code.com/vb/scripts/ShowZip" + zipname, frmSavePathDialog.Text1.Text + nNode.Text + Format(u) + "\" + nNode.Text + ".zip"
                    If FileLen(FilePath + "\" + nNode.Text + ".zip") = 0 Then
                        Kill FilePath + "\" + nNode.Text + ".zip"
                    Else
                        'Dim Zip As New ZipFileClass
                        'Zip.Filename = FilePath + "\" + nNode.Text + ".zip"
                        'Zip.UseDirectoryInfo = True
                        'Zip.Read
                        'ca$ = CurDir$
                        ChDir FilePath
                        'Zip.ExtractDir = frmSavePathDialog.Text1.Text + nNode.Text
                        'Zip.Extract Zip.vFiles
                        Dim na&, nb&
                        VBUnzip FilePath + "\" + nNode.Text + ".zip", CurDir, 0, 1, 0, 1, na, nb

                        ChDir ca$
                    End If
                    Open FilePath + "\" + nNode.Text + ".txt" For Output As #1
                        Print #1, TreeView1.Nodes("DESCR" + ID).Text
                    Close #1
                Else
                    Open frmSavePathDialog.Text1.Text + nNode.Text + ".htm" For Output As #1
                        Print #1, a$
                    Close #1
                End If
                nNode.Checked = False
                nNode.ForeColor = &H80000013
                Open App.Path + "\Downloaded.dat" For Append As #1
                    Print #1, ID
                Close #1
            End If
NextItem:
            Progress.Percent = ((i + 1) / (NumNodi)) * 100
            i = i + 1
        End If
    Next nNode
    frmMain.MousePointer = 0
    Downloading = False
    MsgBox "Download completed!"
End Sub

Private Sub Form_Load()
    Icon = Image1(0).Picture
    Set Tray = New TrayIcon
    Set Tray.OwnerForm = Me
    Set Tray.Icon = Me.Icon
    Tray.Tooltip = "PSC Desktop"
    
    Width = 10230
    PageNumber = "1"
    lblStatus.Caption = lblStatus.Caption & " No url entered to extract html source."
    'WindowState = 1
    myIni.File_Name = App.Path + "\DenzoPSC.ini"
    If Not myIni.FindKeyInSection("General", "SavePath") Then
        Dim BrowseFolder As New CBrowseFolder
        With BrowseFolder
            .hwnd = Me.hwnd
            .Prompt = "Select the default saving destination folder:"
            .ShowStatus = True
            .Folder = App.Path
            .IncludeFiles = False
            .BrowseForFSDirs = True
            .BrowseForFSAncestor = False
            .BrowseForPrinter = False
            .StartUpPosition = vbStartUpOwner
            .Show
            b$ = .Folder + "\"
        End With
        Set BrowseFolder = Nothing
        myIni.WriteData "General", "SavePath", b$
    End If
    
    
    Me.Show
    WebBrowser1.Navigate ("http://www.planet-source-code.com/xq/ASP/lngWId.1/qx/vb/default.htm")
    SavePath.Text = GetSetting("PSC Desktop", "User Defaults", "Save Path", "c:\")
    Set Progress.pic1 = Picture1
    Progress.AddColor vbBlack
    Progress.AddColor vbWhite
    Progress.AddColor vbBlue
    Progress.AddColor vbBlack
    Progress.AddColor vbRed
    Progress.AddColor vbYellow
    Progress.AddColor vbWhite
    Progress.AddColor vbBlack
    Progress.AddColor vbCyan
    Progress.AddColor vbGreen
    Progress.AddColor vbBlack
    Progress.AddColor vbRed
    Progress.AddColor vbYellow
    Progress.AddColor vbGreen
    Progress.AddColor vbBlack
    Progress.AddColor vbWhite
    mnuNewest_Click
End Sub


Private Sub Search_Click()
    Static OldSearch As String
    PathUrl$ = "http://www.planet-source-code.com/xq/ASP/lngWId.1/grpCategories./txtMaxNumberOfEntriesPerPage.100/optSort.Alphabetical/chkThoroughSearch./blnTopCode.False/blnNewestCode.False/blnAuthorSearch.False/lngAuthorId./strAuthorName./blnResetAllVariables./blnEditCode.False/mblnIsSuperAdminAccessOn.False/intFirstRecordOnPage.1/intLastRecordOnPage.10/intMaxNumberOfEntriesPerPage.10/intLastRecordInRecordset.88/chkCodeTypeZip./chkCodeDifficulty./chkCodeTypeText./chkCodeTypeArticle./txtCriteria.%Query%/cmdGoToPage.%PageNumber%/lngMaxNumberOfEntriesPerPage.10/lngWId.1/qx/vb/scripts/BrowseCategoryOrSearchResults.htm"
    a$ = InputBox("Insert words to search for!", "PSC Desktop", OldSearch)
    OldSearch = a$
    a$ = Trim(a$)
    If Trim(a$) = "" Then Exit Sub
    a$ = Replace(a$, " ", "+")
    PageNumber = "1"
    PathUrl$ = Replace(PathUrl$, "%Query%", a$)
    txtURL.Text = PathUrl$
    QueryType$ = "Search: " + a$
    cmdExtract_Click
    cmdLoadTree_Click

End Sub

Private Sub Show_Click()
    WindowState = 0
    Me.Show
End Sub

Private Sub Timer1_Timer()
    Static counter As Integer
    counter = counter + 1
    If Downloading = True Then
        counter = 0
    Else
        If counter > 20 Then
            '10 minutes
            counter = 0
            Dim nod As Node
            Set nod = TreeView1.Nodes("Root").Child
            a$ = ""
            If Not nod Is Nothing Then
                a$ = nod.Text
            End If
            cmdExtract_Click
            cmdLoadTree_Click
            Set nod = TreeView1.Nodes("Root").Child
            If Not nod Is Nothing Then
                If a$ <> nod.Text Then
                    Icon = Image1(1).Picture
                    Set Tray.Icon = Me.Icon
                End If
            End If
        End If
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
        mnuRefresh_Click
      Case 3
         Save_Click
      Case 5
         cmdPrevpage_Click
      Case 6
         cmdNextPage_Click
      Case 8
        mnuBrowser_Click
      Case 9
        Search_Click
      Case 10
        Libraries_Click
      Case 12
        mnuExit_Click
   End Select

End Sub

Private Sub Tray_MouseDblClick(ByVal Button As Integer)
    PopupMenu TrayMenu
End Sub

Private Sub Tray_MouseDown(ByVal Button As Integer)
    PopupMenu TrayMenu
End Sub

Private Sub TreeView1_DblClick()
    Dim nNode As Node
    Set nNode = TreeView1.SelectedItem
    On Error Resume Next
    If Mid$(nNode.Key, 1, 2) = "ID" Then
        If Err > 0 Then Exit Sub
        lblStatus.Caption = "Loading page..."
        WebBrowser1.Navigate (TreeView1.Nodes("URL" + Mid$(nNode.Key, 3)))
    End If
End Sub

Private Sub txtURL_Change()
    'Insert current Page number!
    txtURL.Text = Replace(txtURL.Text, "%PageNumber%", PageNumber)
    Caption = "PSC Desktop " '+ txtURL.Text
End Sub

Private Sub SelectText(ByRef textObj As RichTextBox)
'Selects all the text
    textObj.SelStart = 0
    textObj.SelLength = Len(textObj)
End Sub

Private Sub txtURL_GotFocus()
'Calls the SelectText sub
    Call SelectText(txtURL)
End Sub

Private Sub GetZip(URL As String, ByVal OutName As String)
    'Dim bData() As Byte
    frmMain.MousePointer = 11
    llRetVal = URLDownloadToFile(0, URL, OutName, 0, 0)
    
    'bData() = intMain.OpenURL(URL, icByteArray)
    'Do While intMain.StillExecuting
    'Loop
    'If UBound(bData) > 10 Then
    '    Open OutName For Binary Access Write As #1
    '        Put #1, , bData()
    '    Close #1
    'End If
'    Open OutName For Binary Access Write As #1
'        Put #1, , bData()
'    Close #1
    frmMain.MousePointer = 0

End Sub

Private Function InstrEnd(ByVal SearchedString, ByVal SearchingString) As Long
    InstrEnd = InStr(SearchedString, SearchingString) + Len(SearchingString)
End Function
Private Function GetInsideString(ByVal SearchedString, ByVal PrevString, ByVal SuccString) As String
    On Error Resume Next
    GetInsideString = Mid$(SearchedString, InstrEnd(SearchedString, PrevString), InStr(InstrEnd(SearchedString, PrevString), SearchedString, SuccString) - InstrEnd(SearchedString, PrevString))
End Function
Private Function ClearSpecialChar(ByVal GivenString) As String
    Dim a$
    For i = 1 To Len(GivenString)
        b = Asc(Mid$(GivenString, i, 1))
        If b <> 13 And b <> 10 And b <> 9 Then
            a$ = a$ + Chr$(b)
        End If
    Next i
    ClearSpecialChar = a$
End Function

Private Sub UseSavingKey_Click()
    UseSavingKey.Checked = Not UseSavingKey.Checked
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    DownloadCompleted = True

End Sub

  
  
     

Public Function MakeProper(StringIn As Variant) As String
    'Upper-Cases the first letter of each wo
    '     rd in in a string
    'On Error GoTo HandleErr
    Dim strBuild As String
    Dim intLength As Integer
    Dim intCounter As Integer
    Dim strChar As String
    Dim strPrevChar As String
    intLength = Len(StringIn)
    'Bail out if there is nothing there


    If intLength > 0 Then
        strBuild = UCase(Left(StringIn, 1))


        For intCounter = 1 To intLength
            strPrevChar = Mid$(StringIn, intCounter, 1)
            strChar = Mid$(StringIn, intCounter + 1, 1)


            Select Case strPrevChar
                Case Is = " ", ".", "/"
                strChar = UCase(strChar)
                Case Else
            End Select
        strBuild = strBuild & strChar
        Next intCounter
        MakeProper = strBuild
        'strBuild = MakeWordsLowerCase(strBuild, " and ", " or ", " the ", " a ", " To ")
        MakeProper = strBuild
    End If

End Function

 
 
 
 
 
