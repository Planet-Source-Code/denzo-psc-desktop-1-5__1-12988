Attribute VB_Name = "MBrowseFolder"
Option Explicit

'------------------------------------------------------------------
'Type: TBrowseFolderParams
'------------------------------------------------------------------
'Used internally to set the browse dialog's initial settings.
'------------------------------------------------------------------
Public Type TBrowseFolderParams
    StartUpPosition As StartUpPositionConstants
    nTop As Long
    nLeft As Long
    sCaption As String
    sInitialFolder As String
End Type


'------------------------------------------------------------------
'Function: BrowseCallbackProc
'------------------------------------------------------------------
'The browse dialog box calls this function to notify the application
'about events.
'------------------------------------------------------------------
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal pData As Long) As Long
Attribute BrowseCallbackProc.VB_Description = "The browse dialog box calls this function to notify it about events."
    Dim sDir As String
    Dim sText As String
    Dim tParams As TBrowseFolderParams
    Dim pos As StartUpPositionConstants
    Dim x As Long, y As Long, cx As Long, cy As Long
    Dim scx As Long, scy As Long
    Dim rc As RECT
    Dim hParent As Long
    
    Select Case uMsg
    Case BFFM_INITIALIZED
        CopyMemory tParams, ByVal pData, Len(tParams)
        
        'Change window's title bar text
        sText = tParams.sCaption
        If sText <> "" Then
            SetWindowText hWnd, sText
        End If
        
        'Set initial folder
        sDir = tParams.sInitialFolder
        If sDir <> "" Then
            'WParam is TRUE since you are passing a path.
            'It would be FALSE if you were passing a pidl.
            SendMessageStr hWnd, BFFM_SETSELECTION, True, sDir
        End If
        
        'Set dialog's start-up position
        pos = tParams.StartUpPosition
        
        If pos <> vbStartUpWindowsDefault Then
            GetWindowRect hWnd, rc
            cx = (rc.Right - rc.Left)
            cy = rc.bottom - rc.Top
            
            Select Case pos
            Case vbStartUpManual
                x = tParams.nLeft
                y = tParams.nTop
                MoveWindow ByVal hWnd, ByVal x, ByVal y, ByVal cx, ByVal cy, ByVal 0
            
            Case vbStartUpOwner
                hParent = GetWindowLong(hWnd, GWL_HWNDPARENT)
                
                If hParent <> 0 Then
                    GetWindowRect hParent, rc
                    x = rc.Left + ((rc.Right - rc.Left) - cx) \ 2
                    y = rc.Top + ((rc.bottom - rc.Top) - cy) \ 2
                    
                    'Keep the dialog box within the screen.
                    If x < 0 Then x = 0
                    If y < 0 Then y = 0
                    
                    scx = GetSystemMetrics(SM_CXFULLSCREEN)
                    If (x + cx) > scx Then x = scx - cx
                    
                    scy = GetSystemMetrics(SM_CYFULLSCREEN)
                    If (y + cy) > scy Then y = scy - cy
                    
                    MoveWindow ByVal hWnd, ByVal x, ByVal y, ByVal cx, ByVal cy, ByVal 0
                End If
                
            Case vbStartUpScreen
                x = (GetSystemMetrics(SM_CXFULLSCREEN) - cx) \ 2
                y = (GetSystemMetrics(SM_CYFULLSCREEN) - cy) \ 2
                MoveWindow ByVal hWnd, ByVal x, ByVal y, ByVal cx, ByVal cy, ByVal 0
            End Select
        End If
        
    Case BFFM_SELCHANGED
        sDir = Space$(cMaxPath)
        
        'Set the status window to the currently selected path.
        If SHGetPathFromIDList(lParam, sDir) Then
            SendMessageStr hWnd, BFFM_SETSTATUSTEXT, 0, sDir
        End If
    Case Else
    End Select
    
    BrowseCallbackProc = 0
End Function


