Attribute VB_Name = "SysTray"
Public Const WM_RBUTTONUP = &H205

Global Const WM_MOUSEMOVE = &H200
Global Const NIM_ADD = 0&
Global Const NIM_DELETE = 2&
Global Const NIM_MODIFY = 1&
Global Const NIF_ICON = 2&
Global Const NIF_MESSAGE = 1&
Global Const ABM_GETTASKBARPOS = &H5&

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type

Global Notify As NOTIFYICONDATA
Global BarData As APPBARDATA

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Sub modIcon(form1 As Form, IconID As Long, Icon As Object, Tooltip As String)
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hwnd = form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = Tooltip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)
End Sub

Sub AddIcon(form1 As Form, IconID As Long, Icon As Object, Tooltip As String)
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hwnd = form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = Tooltip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub

Sub delIcon(IconID As Long)
    Dim Result As Long
    Notify.uID = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)
End Sub
Public Sub showTrayIcon(form1 As Form, ByVal Tooltip As String)
    'Set IconObject = form1.Icon
    AddIcon form1, form1.Icon.Handle, form1.Icon, Tooltip

End Sub
Public Sub hideTrayIcon(form1 As Form)
    'delIcon form1.Icon.Handle
    delIcon form1.Icon.Handle
End Sub

