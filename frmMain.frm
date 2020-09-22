VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7560
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrReset 
      Interval        =   20000
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   15
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   12
      Top             =   2280
      Width           =   7530
      Begin Crystal.CrystalReport crpReport 
         Left            =   4200
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3480
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSComctlLib.ImageList imgListTray 
         Left            =   2760
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":08CA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUploadSpeedAverage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average upload speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5730
         TabIndex        =   19
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label lblDownloadSpeedAverage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average download speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   5520
         TabIndex        =   18
         Top             =   120
         Width           =   1890
      End
      Begin VB.Label lblUploadSpeedTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top upload speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblDownloadSpeedTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top download speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1545
      End
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":11A4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.ComboBox cboConnectionType 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   330
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.CommandButton cmdConnectionType 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   345
   End
   Begin VB.Label lblBrokenArrow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broken Arrow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6480
      MouseIcon       =   "frmMain.frx":1A6E
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bandwidth Monitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   495
      TabIndex        =   10
      Top             =   0
      Width           =   6705
   End
   Begin VB.Label lblUSpeed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   4785
   End
   Begin VB.Label lblDSpeed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H00C0FFC0&
      Height          =   345
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   4785
   End
   Begin VB.Label lblSent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   4785
   End
   Begin VB.Label lblRecv 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00C0FFC0&
      Height          =   345
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   4785
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "TokenRing "
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   4425
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Upload speed"
      Height          =   345
      Left            =   15
      TabIndex        =   7
      Top             =   1920
      Width           =   2835
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Download speed"
      Height          =   345
      Left            =   15
      TabIndex        =   6
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Sent bytes"
      Height          =   345
      Left            =   15
      TabIndex        =   5
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Received bytes"
      Height          =   345
      Left            =   15
      TabIndex        =   4
      Top             =   840
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Connection type"
      Height          =   345
      Left            =   15
      TabIndex        =   3
      Top             =   480
      Width           =   2835
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuPopupAlwaysOnTop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnuSystemTrayIconTypeDigital 
         Caption         =   "Digital tray icon"
      End
      Begin VB.Menu mnuSystemTrayIconTypeAnalog 
         Caption         =   "Analog tray icon"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLogToDatabase 
         Caption         =   "Log to database"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report..."
      End
      Begin VB.Menu MBAR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private m_objIpHelper As CIpHelper
Private TransferRate As Single
Private TransferRate2 As Single

Private LastMoment As Date, LastRecvBytes As Long, LastSentBytes As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_MOUSEMOVE = &H200

Private Rcv(1 To 85) As Double
Private Sent(1 To 85) As Double
Private DownloadSpeedTop As Double, UploadSpeedTop As Double, DownloadSpeedAverage As Double, UploadSpeedAverage As Double
Private LoggingInterval As Long, LastLogged As Date

Private Sub cboConnectionType_Click()
cboConnectionType.Visible = False
End Sub

Private Sub cmdConnectionType_Click()
cboConnectionType.Visible = True
End Sub

Private Sub Form_Load()
LastMoment = Now
LastLogged = Now
LoggingInterval = 60

Set m_objIpHelper = New CIpHelper

Dim a As Long
For a = 1 To m_objIpHelper.Interfaces.Count
    cboConnectionType.AddItem m_objIpHelper.Interfaces(a).InterfaceDescription & " "
Next

If Val(GetSetting(App.Title, "Setting", "Connection", 0)) + 1 <= cboConnectionType.ListCount Then
    cboConnectionType.ListIndex = Val(GetSetting(App.Title, "Setting", "Connection", 0))
Else
    cboConnectionType.ListIndex = 0
End If

Me.Move GetSetting(App.Title, "Setting", "Left", Screen.Width - Me.Width), GetSetting(App.Title, "Setting", "Top", Screen.Height - Me.Height - 450)

mnuSystemTrayIconTypeAnalog.Checked = CBool(GetSetting(App.Title, "Setting", "System tray icon type", True))
mnuSystemTrayIconTypeDigital.Checked = Not mnuSystemTrayIconTypeAnalog.Checked
mnuLogToDatabase.Checked = CBool(GetSetting(App.Title, "Setting", "Log to database", False))

If GetSetting(App.Title, "Setting", "Tray", False) = True Then mnuPopupHide_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteIcon

SaveSetting App.Title, "Setting", "Connection", cboConnectionType.ListIndex
SaveSetting App.Title, "Setting", "Left", Me.Left
SaveSetting App.Title, "Setting", "Top", Me.Top
If Not Me.Visible Then SaveSetting App.Title, "Setting", "Tray", True Else SaveSetting App.Title, "Setting", "Tray", False
SaveSetting App.Title, "Setting", "System tray icon type", CBool(mnuSystemTrayIconTypeAnalog.Checked)
SaveSetting App.Title, "Setting", "Log to database", CBool(mnuLogToDatabase.Checked)
End Sub

Private Sub lblBrokenArrow_Click()
ShellExecute hwnd, "open", "mailto:Joy@BDSource.com?Subject=About Bandwidth Monitor", vbNullString, vbNullString, 5
End Sub

Private Sub lblBrokenArrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

lblBrokenArrow.Move lblBrokenArrow.Left + 1, lblBrokenArrow.Top + 1
End Sub

Private Sub lblBrokenArrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X / Screen.TwipsPerPixelX > 4 And X / Screen.TwipsPerPixelX < lblBrokenArrow.Width - 4 And Y / Screen.TwipsPerPixelY > 4 And Y / Screen.TwipsPerPixelY < lblBrokenArrow.Height - 4 Then lblBrokenArrow.Font.Underline = True Else lblBrokenArrow.Font.Underline = False
End Sub

Private Sub lblBrokenArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

lblBrokenArrow.Move lblBrokenArrow.Left - 1, lblBrokenArrow.Top - 1
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuLogToDatabase_Click()
mnuLogToDatabase.Checked = Not mnuLogToDatabase.Checked
End Sub

Private Sub mnuPopupAlwaysOnTop_Click()
mnuPopupAlwaysOnTop.Checked = Not mnuPopupAlwaysOnTop.Checked
If mnuPopupAlwaysOnTop.Checked Then
    StayOnTop Me, True
Else
    StayOnTop Me, False
End If
End Sub

Private Sub mnuPopupExit_Click()
Unload Me
End Sub

Private Sub mnuPopupHide_Click()
CreateIcon
Me.Hide
End Sub

Private Sub mnuReport_Click()
crpReport.DataFiles(0) = App.Path & "\BM.mdb"
crpReport.ReportFileName = App.Path & "\BM.rpt"
crpReport.Action = 1
End Sub

Private Sub mnuSystemTrayIconTypeAnalog_Click()
mnuSystemTrayIconTypeDigital.Checked = mnuSystemTrayIconTypeAnalog.Checked
mnuSystemTrayIconTypeAnalog.Checked = Not mnuSystemTrayIconTypeAnalog.Checked
End Sub

Private Sub mnuSystemTrayIconTypeDigital_Click()
mnuSystemTrayIconTypeAnalog.Checked = mnuSystemTrayIconTypeDigital.Checked
mnuSystemTrayIconTypeDigital.Checked = Not mnuSystemTrayIconTypeDigital.Checked
End Sub

Private Sub picTray_Click()
PopupMenu mnuPopup
End Sub

Private Sub picTray_DblClick()
Unload Me
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible Then Exit Sub
    Select Case X / Screen.TwipsPerPixelX
    Case Is = WM_LBUTTONDOWN
        Me.Show
        DeleteIcon
    Case Is = WM_RBUTTONDOWN
    Case Is = WM_MOUSEMOVE
    End Select
End Sub

Private Sub tmrReset_Timer()
DownloadSpeedTop = 0
UploadSpeedTop = 0
'DownloadSpeedAverage = 0
'UploadSpeedAverage = 0
End Sub

Private Sub tmrUpdate_Timer()
On Error Resume Next

If DateDiff("s", LastMoment, Now) < 1 Then Exit Sub

tmrUpdate.Enabled = False
    
Dim objInterface As CInterface
Set objInterface = m_objIpHelper.Interfaces(cboConnectionType.ListIndex + 1)

lblType = m_objIpHelper.Interfaces(cboConnectionType.ListIndex + 1).InterfaceDescription & " "

Dim BytesRecv As Long, BytesSent As Long
BytesRecv = m_objIpHelper.BytesReceived
BytesSent = m_objIpHelper.BytesSent

lblRecv.Caption = Format(BytesRecv / 1024, "###,###,###,###,##0 KB")
lblSent.Caption = Format(BytesSent / 1024, "###,###,###,###,##0 KB")

Dim DS As Long, US As Long
DS = BytesRecv - LastRecvBytes
US = BytesSent - LastSentBytes
If DownloadSpeedTop < DS Then
    tmrReset.Enabled = False
    tmrReset.Enabled = True
    DownloadSpeedTop = DS
End If
If UploadSpeedTop < US Then
    tmrReset.Enabled = False
    tmrReset.Enabled = True
    UploadSpeedTop = US
End If
DownloadSpeedAverage = (DownloadSpeedAverage + DS) / 2
UploadSpeedAverage = (UploadSpeedAverage + US) / 2
lblDownloadSpeedTop = "Top download speed: " & Format(DownloadSpeedTop / 1024, "###,###,###,###,#0.#0 Kb/S")
lblUploadSpeedTop = "Top upload speed: " & Format(UploadSpeedTop / 1024, "###,###,###,###,#0.#0 Kb/S")
lblDownloadSpeedAverage = "Average download speed: " & Format(DownloadSpeedAverage / 1024, "###,###,###,###,#0.#0 Kb/S")
lblUploadSpeedAverage = "Average upload speed: " & Format(UploadSpeedAverage / 1024, "###,###,###,###,#0.#0 Kb/S")

If DS / 1024 < 1 Then
    lblDSpeed = Format(DS, "0 BS ")
Else
    lblDSpeed = Format(DS / 1024, "0.#0 KBS ")
End If
If US / 1024 < 1 Then
    lblUSpeed = Format(US, "0 BS ")
Else
    lblUSpeed = Format(US / 1024, "0.#0 KBS ")
End If

UpdateGraph DS, US

LastRecvBytes = BytesRecv
LastSentBytes = BytesSent
LastMoment = Now

If m_objIpHelper.Interfaces.Count <> cboConnectionType.ListCount Then
    Dim a As Long
    cboConnectionType.Clear
    For a = 1 To m_objIpHelper.Interfaces.Count
        cboConnectionType.AddItem m_objIpHelper.Interfaces(a).InterfaceDescription & " "
    Next
    If Val(GetSetting(App.Title, "Setting", "Connection", 0)) + 1 <= cboConnectionType.ListCount Then
        cboConnectionType.ListIndex = Val(GetSetting(App.Title, "Setting", "Connection", 0))
    Else
        cboConnectionType.ListIndex = 0
    End If
End If

Log2DB DS, US

tmrUpdate.Enabled = True
End Sub

Private Sub UpdateGraph(NewRcv As Long, NewSent As Long)
On Error Resume Next

Dim a As Long, TopRcv As Double, TopSent As Double, vTop As Double, Frq As Long

Frq = 85

For a = 2 To Frq
    Rcv(a - 1) = Rcv(a)
    Sent(a - 1) = Sent(a)
    
    If Rcv(a) > TopRcv Then TopRcv = Rcv(a)
    If Sent(a) > TopSent Then TopSent = Sent(a)
Next
Rcv(Frq) = NewRcv
Sent(Frq) = NewSent

If Rcv(Frq) > TopRcv Then TopRcv = Rcv(Frq)
If Sent(Frq) > TopSent Then TopSent = Sent(Frq)

If TopRcv > TopSent Then vTop = TopRcv Else vTop = TopSent

picGraph.Cls

If Me.Visible Then
    If picGraph.BackColor = vbBlack Then picGraph.BackColor = vbWhite
    picGraph.PSet (13, 1), vbWhite
    picGraph.ForeColor = &HE0E0E0
    picGraph.Print "Joy Softwares"
    picGraph.PSet (11, -1), vbWhite
    picGraph.ForeColor = &HFFEFEF
    picGraph.Print "Joy Softwares"
End If

For a = 1 To Frq
    picGraph.Line ((a - 1) * (picGraph.ScaleWidth / Frq), picGraph.ScaleHeight - 1)-(a * (picGraph.ScaleWidth / Frq) - 1, picGraph.ScaleHeight - (picGraph.ScaleHeight * (Rcv(a) / vTop)) - 1), RGB(0, 255, 0), BF
    picGraph.Line ((a - 1) * (picGraph.ScaleWidth / Frq), picGraph.ScaleHeight - 1)-(a * (picGraph.ScaleWidth / Frq) - 1, picGraph.ScaleHeight - (picGraph.ScaleHeight * (Sent(a) / vTop)) - 1), RGB(255, 0, 0), BF
Next

If mnuSystemTrayIconTypeAnalog.Checked = True Then
    picIcon.PaintPicture picGraph.Image, 0, 0, picIcon.ScaleWidth, picIcon.ScaleHeight, picGraph.ScaleWidth - picGraph.ScaleHeight, 0, picGraph.ScaleHeight, picGraph.ScaleHeight
Else
    picIcon.Cls
    If TextWidth(Format(NewRcv / 1024, "##0.0")) > picIcon.ScaleWidth Then picIcon.PSet (0, -4) Else picIcon.PSet ((picIcon.ScaleWidth - TextWidth(Format(NewRcv / 1024, "##0.0"))) / 2, -4)
    picIcon.ForeColor = RGB(0, 255, 0)
    picIcon.Print Format(NewRcv / 1024, "##0.0")
    If TextWidth(Format(NewSent / 1024, "##0.0")) > picIcon.ScaleWidth Then picIcon.PSet (0, picIcon.ScaleHeight / 2 - 4) Else picIcon.PSet ((picIcon.ScaleWidth - TextWidth(Format(NewSent / 1024, "##0.0"))) / 2, picIcon.ScaleHeight / 2 - 4)
    picIcon.ForeColor = RGB(255, 150, 150)
    picIcon.Print Format(NewSent / 1024, "##0.0")
End If

If Not Me.Visible Then
    If picGraph.BackColor = vbWhite Then picGraph.BackColor = vbBlack
    imgListTray.ListImages.Remove 1
    imgListTray.ListImages.Add , , picIcon.Image
    ModifyIcon
End If
End Sub

Sub Log2DB(DownloadSpeed As Long, UploadSpeed As Long)
If DateDiff("s", LastLogged, Now) < LoggingInterval Or mnuLogToDatabase.Checked = False Then Exit Sub

OpenDatabase(App.Path & "\BM.mdb").Execute "INSERT INTO tblLog (LogDate, LogTime, DownLoadSpeed, UploadSpeed) VALUES (#" & Date & "#, #" & Time & "#, " & DownloadSpeed & ", " & UploadSpeed & ")"
LastLogged = Now
End Sub
