Attribute VB_Name = "Main"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 19 'Replace the szTip string's length with your tip's length
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const WM_MOUSEMOVE = &H200

Public Sub DragForm(frm As Form)
On Local Error Resume Next

ReleaseCapture
SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Public Sub StayOnTop(frm As Form, OnTop As Boolean)
If OnTop Then
    SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Else
    SetWindowPos frm.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Public Sub CreateIcon() 'Call this method to create the tray icon
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = frmMain.picTray.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frmMain.picTray.Picture
    Tic.szTip = "System Tray Example"
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub ModifyIcon() 'Call this method to modify the trat icon properties
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = frmMain.picTray.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frmMain.imgListTray.ListImages(1).ExtractIcon
    Tic.szTip = "System Tray Example"
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub

Public Sub DeleteIcon() 'Call this method to remove the tray icon
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = frmMain.picTray.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

