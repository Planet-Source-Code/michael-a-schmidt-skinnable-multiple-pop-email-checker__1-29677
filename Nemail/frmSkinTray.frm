VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSkinTray 
   BorderStyle     =   0  'None
   ClientHeight    =   2880
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   12285
   Icon            =   "frmSkinTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer timMinutes 
      Interval        =   60000
      Left            =   4530
      Top             =   630
   End
   Begin VB.PictureBox picControls 
      Align           =   1  'Align Top
      BackColor       =   &H00C0E0FF&
      Height          =   2835
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   12225
      TabIndex        =   0
      Top             =   0
      Width           =   12285
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4470
         Top             =   180
      End
      Begin VB.PictureBox picMail 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   6780
         Picture         =   "frmSkinTray.frx":0CCA
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   390
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picMail 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   6510
         Picture         =   "frmSkinTray.frx":0E14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMail 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   5970
         Picture         =   "frmSkinTray.frx":0F5E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   150
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   3300
         Top             =   105
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
               Picture         =   "frmSkinTray.frx":1828
               Key             =   "Exit"
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   3930
         Top             =   150
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Image imgSkins 
         Height          =   2760
         Index           =   12
         Left            =   7590
         Picture         =   "frmSkinTray.frx":2120
         Top             =   30
         Width           =   4590
      End
      Begin VB.Image imgSkins 
         Height          =   330
         Index           =   0
         Left            =   30
         Picture         =   "frmSkinTray.frx":5E7D
         Top             =   15
         Width           =   1425
      End
      Begin VB.Image imgSkins 
         Height          =   330
         Index           =   1
         Left            =   30
         Picture         =   "frmSkinTray.frx":6364
         Top             =   390
         Width           =   1425
      End
      Begin VB.Image imgSkins 
         Height          =   165
         Index           =   2
         Left            =   1530
         Picture         =   "frmSkinTray.frx":67D4
         Top             =   75
         Width           =   165
      End
      Begin VB.Image imgSkins 
         Height          =   165
         Index           =   3
         Left            =   1515
         Picture         =   "frmSkinTray.frx":6B72
         Top             =   450
         Width           =   165
      End
      Begin VB.Image imgSkins 
         Height          =   225
         Index           =   4
         Left            =   1860
         Picture         =   "frmSkinTray.frx":6EC6
         Top             =   90
         Width           =   225
      End
      Begin VB.Image imgSkins 
         Height          =   225
         Index           =   5
         Left            =   1860
         Picture         =   "frmSkinTray.frx":72F6
         Top             =   465
         Width           =   240
      End
      Begin VB.Image imgSkins 
         Height          =   225
         Index           =   6
         Left            =   2250
         Picture         =   "frmSkinTray.frx":773A
         Top             =   105
         Width           =   225
      End
      Begin VB.Image imgSkins 
         Height          =   225
         Index           =   7
         Left            =   2235
         Picture         =   "frmSkinTray.frx":7B3E
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgSkins 
         Height          =   195
         Index           =   8
         Left            =   2625
         Picture         =   "frmSkinTray.frx":7F5A
         Top             =   75
         Width           =   195
      End
      Begin VB.Image imgSkins 
         Height          =   195
         Index           =   9
         Left            =   2625
         Picture         =   "frmSkinTray.frx":8383
         Top             =   450
         Width           =   195
      End
      Begin VB.Image imgSkins 
         Height          =   225
         Index           =   10
         Left            =   2970
         Picture         =   "frmSkinTray.frx":879D
         Top             =   90
         Width           =   225
      End
      Begin VB.Image imgSkins 
         Height          =   195
         Index           =   11
         Left            =   2970
         Picture         =   "frmSkinTray.frx":8BCD
         Top             =   465
         Width           =   195
      End
      Begin VB.Image imgSkins 
         Height          =   4500
         Index           =   13
         Left            =   0
         Picture         =   "frmSkinTray.frx":8FE7
         Top             =   0
         Width           =   7500
      End
   End
   Begin VB.Menu mnuMainMenu 
      Caption         =   "MainMenu"
      Begin VB.Menu mnuMailbox 
         Caption         =   "Mailbox"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckMail 
         Caption         =   "Check Mail"
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSkinTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MinutesElapsed As Integer   ' - Mail Check Timer


Private Sub Form_Load()

    On Local Error Resume Next

    Call LoadINISettings
    Call LoadSkins(Me, True)
    Call SetColors(Me)

    'Set visible to true so that the main menu will be visible with the login dialog box in front of it...
    Me.Visible = True
    DoEvents

    'Show the login screen...
    Login.ReLoggingIn = False

    Timer1.Enabled = True
    Me.Hide

    MailTray.ShowIcon Me
    MailTray.ChangeIcon Me, picMail.Item(0)

End Sub
Sub LoadINISettings()

'Form properties...
If Trim$(ReadINI(Me.Name, "Caption", QuickRef.GlobalINIFileName)) <> "" Then
    Me.Caption = ReadINI(Me.Name, "Caption", QuickRef.GlobalINIFileName)
Else
    Me.Caption = "Nemail 2.0"
End If

'Form Coordinates...
Me.WindowState = Val(ReadINI(Me.Name, "WindowState", QuickRef.UserINIFileName))
If Me.WindowState = vbMaximized Then Exit Sub
Me.Left = Val(ReadINI(Me.Name, "Left", QuickRef.UserINIFileName))
Me.Top = Val(ReadINI(Me.Name, "Top", QuickRef.UserINIFileName))
Me.Height = Val(ReadINI(Me.Name, "Height", QuickRef.UserINIFileName))
Me.Width = Val(ReadINI(Me.Name, "Width", QuickRef.UserINIFileName))

End Sub
Private Sub Form_Resize()

On Local Error Resume Next

'MainMenu PictureBox...
'picMainMenu.Height = Me.Height - 1640

'Panels...
'picMainMenuPanel.Top = ((Me.Height / 2 - picMainMenuPanel.Height / 2) - Toolbar1.Height - 200)
'picMainMenuPanel.Left = (Me.Width / 2 - picMainMenuPanel.Width / 2) - 55

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save this form's settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'WindowState...
Call WriteINI(Me.Name, "WindowState", Me.WindowState, QuickRef.UserINIFileName)

'If Window State = vbNormal Then...
If Me.WindowState = vbNormal Then
    Call WriteINI(Me.Name, "Left", Me.Left, QuickRef.UserINIFileName)
    Call WriteINI(Me.Name, "Top", Me.Top, QuickRef.UserINIFileName)
    Call WriteINI(Me.Name, "Height", Me.Height, QuickRef.UserINIFileName)
    Call WriteINI(Me.Name, "Width", Me.Width, QuickRef.UserINIFileName)
End If

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuCheckMail_Click()
    CheckArrayMail
End Sub

Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuMailbox_Click()
    frmMain.WindowState = vbNormal
    frmMain.Show
End Sub

Private Sub mnuSetup_Click()
    frmMain.Show
    frmMain.WindowState = vbNormal
    frmColors.Show
    frmColors.WindowState = vbNormal
End Sub

Private Sub mnuToolTip_Click()
    ViewToolTip
End Sub

Private Sub Timer1_Timer()

On Local Error Resume Next

Dim x As Byte
Dim iAutoApply As Boolean
Dim iShowHideSolidColors As Boolean
Dim iWindowState As Byte

'Update Colors...
If Colors.UpdateColors = True Then
    Call LoadColors
    For x = 0 To Forms.Count - 1
        Call SetColors(Forms(x))
    Next x
    Colors.UpdateColors = False
End If

'Change Skin...
If Skins.UpdateSkins = True Then
    iAutoApply = (frmColors.imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture)
    iShowHideSolidColors = (frmColors.imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckONID).Picture)
    iWindowState = frmSkinTray.WindowState
    
    ' Before loading other skins, load ourself!
    Call LoadSkins(Me, False)
    For x = 0 To Forms.Count - 1
        If Forms(x).Name <> Me.Name Then Call LoadSkins(Forms(x), True)
        If Forms(x).Name <> Me.Name Then Call SetColors(Forms(x))
    Next x
    'frmSkinTray.WindowState = vbMinimized
    'frmSkinTray.WindowState = iWindowState
    Skins.UpdateSkins = False
    Unload frmColors
    DoEvents
    frmColors.Show
    frmColors.ZOrder
    'Auto Apply...
    If iAutoApply Then
        frmColors.imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture
    Else
        frmColors.imgAutoApply.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture
    End If
    'Show / Hide Solid Colors...
    If iShowHideSolidColors Then
        frmColors.imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckONID).Picture
        frmColors.picColorsSquare.Visible = True
    Else
        frmColors.imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture
        frmColors.picColorsSquare.Visible = False
    End If
End If

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo ToolBar1_ButtonClickError

'Toolbar Buttons...
Select Case LCase$(Button.Key)
    'Employee...
    Case "employee"
    'Clients...
    Case "clients"
    'Sign-In...
    Case "signin"
    'Payroll...
    Case "payroll"
    'Billing...
    Case "billing"
    'Colors...
    Case "colors"
        frmColors.Show
        frmColors.WindowState = vbNormal
        frmColors.ZOrder
    'Exit...
    Case "exit"
        Unload Me
End Select

Exit Sub



ToolBar1_ButtonClickError:
    Call WriteToErrorLog(Me.Name, "ToolBar1_ButtonClickError", Err.Description, Err.Number, False)
    Exit Sub

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Remember..... The value of X will be different if the icon is minimized
' to the system tray.  The values in this case will be as follows,
'       7680   ' MouseMove
'       7695   ' Left MouseDown
'       7710   ' Left MouseUp
'       7725   ' Left DoubleClick
'       7740   ' Right MouseDown
'       7755   ' Right MouseUp
'       7770   ' Right DoubleClick
If MailTray.bRunningInTray Then          'Check to see if form is in the system tray

    frmToolTip.Left = Screen.Width - frmToolTip.Width
    frmToolTip.Top = Screen.Height - frmToolTip.Height - 420
    
    Debug.Print x

    Select Case x                           'If it is, use X to get message value
        Case 7755: PopupMenu Me.mnuMainMenu, vbPopupMenuRightButton
        Case 7710: ViewToolTip
        'Case 7755: PopupMenu Me.mnuMailMenu, vbPopupMenuRightButton
        'Case 513: ViewToolTip
        'Case 7725: Me.Show
    End Select
End If

End Sub



Private Sub ViewToolTip()
Static ToolVisible As Boolean
    If ToolVisible Then
        frmToolTip.Hide
        ToolVisible = False
    Else
        frmToolTip.Show
        ToolVisible = True
    End If
End Sub

'================================
'   Timer
'================================
Private Sub timMinutes_Timer()
    '#############################
    '# Every minute this sub is
    '# called, we simply increment
    '# our counter or reset and
    '# check the mail.
    MinutesElapsed = MinutesElapsed + 1

    If MinutesElapsed = frmMain.txtDelay Then
        CheckArrayMail
        MinutesElapsed = 0    ' Reset Counter
    End If

End Sub
