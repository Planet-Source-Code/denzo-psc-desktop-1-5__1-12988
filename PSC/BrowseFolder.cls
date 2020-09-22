VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "BrowseFolder dialog box."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'------------------------------------------------------------------
'Class: CBrowseFolder
'------------------------------------------------------------------
'This class allows user to select folders using the standard
'BrowseForFolder dialog box.
'
'References -> Win.TLB
'
'For bugs reports and suggestions mail to:
'   mauritricoli@hotmail.com
'------------------------------------------------------------------


'------------------------------------------------------------------
'Enum ESpecialFoldersConstant
'------------------------------------------------------------------
'Used to select the root of the BrowseForFolder dialog's treeview.
'------------------------------------------------------------------
Public Enum ESpecialFoldersConstant
    esfAppData = CSIDL_APPDATA
    esfBitBucket = CSIDL_BITBUCKET
    esfCommonDesktopDir = CSIDL_COMMON_DESKTOPDIRECTORY
    esfCommonPrograms = CSIDL_COMMON_PROGRAMS
    esfCommonStartMenu = CSIDL_COMMON_STARTMENU
    esfCommonStartUp = CSIDL_COMMON_STARTUP
    esfControlPanel = CSIDL_CONTROLS
    esfDesktop = CSIDL_DESKTOP
    esfDesktopDir = CSIDL_DESKTOPDIRECTORY
    esfDrives = CSIDL_DRIVES
    esfFavorites = CSIDL_FAVORITES
    esfFonts = CSIDL_FONTS
    esfNethood = CSIDL_NETHOOD
    esfNetwork = CSIDL_NETWORK
    esfPersonal = CSIDL_PERSONAL
    esfPrinters = CSIDL_PRINTERS
    esfPrintHood = CSIDL_PRINTHOOD
    esfPrograms = CSIDL_PROGRAMS
    esfRecent = CSIDL_RECENT
    esfSendTo = CSIDL_SENDTO
    esfStartMenu = CSIDL_STARTMENU
    esfStartUp = CSIDL_STARTUP
    esfTemplates = CSIDL_TEMPLATES
End Enum

'------------------------------------------------------------------
'Properties
'------------------------------------------------------------------
Private hOwnerWnd As Long
Private BrowseOptions As Win.EBROWSEOPTIONS
Private eStartUpPosition As StartUpPositionConstants
Private sFolder As String
Private sPrompt As String
Private sCaption As String
Private nTop As Long
Private nLeft As Long
Private nRoot As ESpecialFoldersConstant

'------------------------------------------------------------------
'Property: BrowseForComputer
'------------------------------------------------------------------
'Only return computers. If the user selects anything other than a
'computer, the OK button is grayed.
'------------------------------------------------------------------
Public Property Let BrowseForComputer(ByVal fValue As Boolean)
Attribute BrowseForComputer.VB_Description = "Only return computers. If the user selects anything other than a computer, the OK button is grayed. "
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_BROWSEFORCOMPUTER)
    bMask = fValue And BIF_BROWSEFORCOMPUTER
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get BrowseForComputer() As Boolean
    BrowseForComputer = (BrowseOptions And BIF_BROWSEFORCOMPUTER) <> 0
End Property


'------------------------------------------------------------------
'Property: BrowseForFSAncestor
'------------------------------------------------------------------
'Only return file system ancestors. If the user selects anything
'other than a file system ancestor, the OK button is grayed.
'------------------------------------------------------------------
Public Property Let BrowseForFSAncestor(ByVal fValue As Boolean)
Attribute BrowseForFSAncestor.VB_Description = "Only return file system ancestors. If the user selects anything other than a file system ancestor, the OK button is grayed."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_RETURNFSANCESTORS)
    bMask = fValue And BIF_RETURNFSANCESTORS
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get BrowseForFSAncestor() As Boolean
    BrowseForFSAncestor = (BrowseOptions And BIF_RETURNFSANCESTORS) <> 0
End Property


'------------------------------------------------------------------
'Property: BrowseForFSDirs
'------------------------------------------------------------------
'Only return file system directories. If the user selects folders
'that are not part of the file system, the OK button is grayed.
'------------------------------------------------------------------
Public Property Let BrowseForFSDirs(ByVal fValue As Boolean)
Attribute BrowseForFSDirs.VB_Description = "Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_RETURNONLYFSDIRS)
    bMask = fValue And BIF_RETURNONLYFSDIRS
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get BrowseForFSDirs() As Boolean
    BrowseForFSDirs = (BrowseOptions And BIF_RETURNONLYFSDIRS) <> 0
End Property


'------------------------------------------------------------------
'Property: BrowseForPrinter
'------------------------------------------------------------------
'Only return printers. If the user selects anything other than a
'printer, the OK button is grayed.
'------------------------------------------------------------------
Public Property Let BrowseForPrinter(ByVal fValue As Boolean)
Attribute BrowseForPrinter.VB_Description = "Only return printers. If the user selects anything other than a printer, the OK button is grayed."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_BROWSEFORPRINTER)
    bMask = fValue And BIF_BROWSEFORPRINTER
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get BrowseForPrinter() As Boolean
    BrowseForPrinter = (BrowseOptions And BIF_BROWSEFORPRINTER) <> 0
End Property


'------------------------------------------------------------------
'Property: Caption
'------------------------------------------------------------------
'The text displayed in the dialog's title bar.
'------------------------------------------------------------------
Public Property Let Caption(ByVal sText As String)
Attribute Caption.VB_Description = "The text displayed in the dialog's title bar."
    sCaption = sText
End Property

Public Property Get Caption() As String
    Caption = sCaption
End Property


'------------------------------------------------------------------
'Property: DontGoBelowDomain
'------------------------------------------------------------------
'Do not include network folders below the domain level in the tree
'view control.
'------------------------------------------------------------------
Public Property Let DontGoBelowDomain(ByVal fValue As Boolean)
Attribute DontGoBelowDomain.VB_Description = "Do not include network folders below the domain level in the tree view control."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_DONTGOBELOWDOMAIN)
    bMask = fValue And BIF_DONTGOBELOWDOMAIN
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get DontGoBelowDomain() As Boolean
    DontGoBelowDomain = (BrowseOptions And BIF_DONTGOBELOWDOMAIN) <> 0
End Property


'------------------------------------------------------------------
'Function: FuncAddr
'------------------------------------------------------------------
'Wrapper function for AddressOf operator (for internal use only).
'------------------------------------------------------------------
Private Function FuncAddr(ByVal pFunc As Long) As Long
    FuncAddr = pFunc
End Function


'------------------------------------------------------------------
'Property: Folder
'------------------------------------------------------------------
'The folder selected by the user, or the default folder.
'------------------------------------------------------------------
Public Property Let Folder(ByVal sInitialDir As String)
Attribute Folder.VB_Description = "The folder selected by the user, or the default folder."
Attribute Folder.VB_UserMemId = 0
    sFolder = sInitialDir
End Property

Public Property Get Folder() As String
    Folder = sFolder
End Property


'------------------------------------------------------------------
'Property: hWnd
'------------------------------------------------------------------
'Handle to the owner window for the dialog box.
'------------------------------------------------------------------
Public Property Let hWnd(ByVal hOwner As Long)
Attribute hWnd.VB_Description = "Handle to the owner window for the dialog box."
    hOwnerWnd = hOwner
End Property

Public Property Get hWnd() As Long
    hWnd = hOwnerWnd
End Property


'------------------------------------------------------------------
'Property: IncludeFiles
'------------------------------------------------------------------
'The browse dialog will display files as well as folders.
'------------------------------------------------------------------
Public Property Let IncludeFiles(ByVal fValue As Boolean)
Attribute IncludeFiles.VB_Description = "If True, the browse dialog will display files as well as folders."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_BROWSEINCLUDEFILES)
    bMask = fValue And BIF_BROWSEINCLUDEFILES
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get IncludeFiles() As Boolean
    IncludeFiles = (BrowseOptions And BIF_BROWSEINCLUDEFILES) <> 0
End Property

'------------------------------------------------------------------
'Property: Left
'------------------------------------------------------------------
'Coordinates of the left edge of the dialog box.
'------------------------------------------------------------------
Public Property Let Left(ByVal x As Long)
Attribute Left.VB_Description = "Coordinates of the left edge of the dialog box."
    nLeft = x
End Property

Public Property Get Left() As Long
    Left = nLeft
End Property


'------------------------------------------------------------------
'Property: Root
'------------------------------------------------------------------
'The root folder for the BrowseForFolder dialog's treeview.
'------------------------------------------------------------------
Public Property Let Root(ByVal SpecialFolder As ESpecialFoldersConstant)
Attribute Root.VB_Description = "The root folder for the BrowseForFolder dialog's treeview."
    nRoot = SpecialFolder
End Property

Public Property Get Root() As ESpecialFoldersConstant
    Root = nRoot
End Property

'------------------------------------------------------------------
'Method: Show
'------------------------------------------------------------------
'Displays a dialog box that enables the user to select a shell
'folder.
'------------------------------------------------------------------
Public Function Show() As Boolean
Attribute Show.VB_Description = "Displays a dialog box that enables the user to select a shell folder."
    Dim tParams As TBrowseFolderParams
    Dim tBrowseInfo As Win.BROWSEINFO
    
    'Initialize custom parameter structure
    ZeroMemory tParams, Len(tParams)
    
    With tParams
        .nLeft = nLeft
        .nTop = nTop
        .sCaption = sCaption
        .sInitialFolder = sFolder
        .StartUpPosition = eStartUpPosition
    End With
    
    'Initialize TBrowseInfo structure.
    ZeroMemory tBrowseInfo, Len(tBrowseInfo)
    
    With tBrowseInfo
        If sPrompt <> "" Then
            .lpszTitle = lstrcat(sPrompt, "")
        End If
        .hwndOwner = hOwnerWnd
        .lpfn = FuncAddr(AddressOf BrowseCallbackProc)
        .ulFlags = BrowseOptions
        .pidlRoot = nRoot
        .lParam = VarPtr(tParams)
    End With
    
    'Get user's selection.
    Dim lpIDList As Long
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        Dim pAlloc As Win.IVBMalloc
        'Get the interface pointer to the shell's task allocator.
        SHGetMalloc pAlloc
        
        Dim sBuffer As String
        sBuffer = String(cMaxPath, 0)
        SHGetPathFromIDList lpIDList, sBuffer
        sFolder = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        
        'Free the returned item identifier list.
        pAlloc.Free lpIDList
        Set pAlloc = Nothing
        Show = True
    Else
        sFolder = ""
        Show = False
    End If
End Function


'------------------------------------------------------------------
'Property: ShowStatus
'------------------------------------------------------------------
'Indicates wether or not the dialog box will include a status area for showing the currently selected folder.
'------------------------------------------------------------------
Public Property Let ShowStatus(ByVal fValue As Boolean)
Attribute ShowStatus.VB_Description = "Indicates wether or not the dialog box will include a status area for showing the currently selected folder."
    Dim bSaved As Long, bMask As Long
    
    bSaved = BrowseOptions And (Not BIF_STATUSTEXT)
    bMask = fValue And BIF_STATUSTEXT
    BrowseOptions = bSaved Xor bMask
End Property

Public Property Get ShowStatus() As Boolean
    ShowStatus = (BrowseOptions And BIF_STATUSTEXT) <> 0
End Property


'------------------------------------------------------------------
'Property: Prompt
'------------------------------------------------------------------
'String that is displayed above the tree view control in the dialog
'box.  This string can be used to specify instructions to the user.
'------------------------------------------------------------------
Public Property Let Prompt(ByVal sText As String)
Attribute Prompt.VB_Description = "String that is displayed above the tree view control in the dialog box.  This string can be used to specify instructions to the user."
    sPrompt = sText
End Property

Public Property Get Prompt() As String
    Prompt = sPrompt
End Property


'------------------------------------------------------------------
'Property: StartUpPosition
'------------------------------------------------------------------
'Returns or sets a value specifying the position of the dialog when
'it first appears.
'------------------------------------------------------------------
Public Property Let StartUpPosition(ByVal Value As StartUpPositionConstants)
Attribute StartUpPosition.VB_Description = "Returns or sets a value specifying the position of the dialog when it first appears."
    eStartUpPosition = Value
End Property

Public Property Get StartUpPosition() As StartUpPositionConstants
    StartUpPosition = eStartUpPosition
End Property


'------------------------------------------------------------------
'Property: Top
'------------------------------------------------------------------
'Coordinates of the top edge of the dialog box.
'------------------------------------------------------------------
Public Property Let Top(ByVal y As Long)
Attribute Top.VB_Description = "Coordinates of the top edge of the dialog box."
    nTop = y
End Property

Public Property Get Top() As Long
    Top = nTop
End Property


Private Sub Class_Initialize()
    nRoot = 0
    BrowseOptions = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
    eStartUpPosition = vbStartUpWindowsDefault
End Sub


