VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Mailbox"
   ClientHeight    =   2715
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4695
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4695
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4110
      TabIndex        =   6
      Text            =   "5"
      ToolTipText     =   "Enter length in minutes to check mail."
      Top             =   450
      Width           =   255
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   1365
      Left            =   2160
      TabIndex        =   0
      Top             =   750
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2408
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgMailIcons"
      SmallIcons      =   "imgMailIcons"
      ColHdrIcons     =   "imgMailIcons"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Server"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   3863
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Mail"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMailIcons 
      Left            =   0
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3240
      MouseIcon       =   "frmMain.frx":15BC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "Label"
      ToolTipText     =   "Click here to modify your account."
      Top             =   480
      Width           =   465
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frmMain.frx":170E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "Label"
      ToolTipText     =   "Click here to delete your account."
      Top             =   480
      Width           =   465
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2160
      MouseIcon       =   "frmMain.frx":1860
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "Label"
      ToolTipText     =   "Click here to add your account."
      Top             =   480
      Width           =   285
   End
   Begin VB.Label lblCheckMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to check your mail."
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label lblSetup 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2010
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here for setup."
      Top             =   2280
      Width           =   435
   End
   Begin VB.Image imgSetup 
      Height          =   345
      Left            =   1590
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here for setup."
      Top             =   2220
      Width           =   1365
   End
   Begin VB.Image imgCheckMail 
      Height          =   345
      Left            =   3000
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to check your mail."
      Top             =   2220
      Width           =   1365
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Left            =   3930
      Tag             =   "Minimize"
      ToolTipText     =   "Minimize"
      Top             =   30
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   4230
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Tag2 As String

Public Property Let Tag2(sValue As String)

m_Tag2 = sValue

End Property
Public Property Get Tag2() As String

Tag2 = m_Tag2

End Property



Private Sub Form_Load()
On Local Error Resume Next

Call LoadINISettings
Call LoadSkins(Me, True)
Call SetColors(Me)

'Form Coordinates...
Me.Width = 4590
Me.Height = 2760

End Sub
Sub LoadINISettings()

'Set the custom Tag2 property (for skins)...
Me.Tag2 = "306x184"

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left", QuickRef.UserINIFileName))
Me.Top = Val(ReadINI(Me.Name, "Top", QuickRef.UserINIFileName))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
Dim mCounter As Integer

    For mCounter = 0 To lblMenu.Count - 1
        lblMenu(mCounter).ForeColor = Colors.LabelForeColor
    Next mCounter

End Sub
Public Sub Form_Unload(Cancel As Integer)

'Save this form's settings...
    UnloadMailArray
    Call SaveINISettings
    MailTray.RemoveIcon frmSkinTray
    Unload Me
    End

End Sub
Sub SaveINISettings()

On Local Error Resume Next

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left, QuickRef.UserINIFileName)
Call WriteINI(Me.Name, "Top", Me.Top, QuickRef.UserINIFileName)

End Sub


Private Sub imgClose_Click()
    Me.Hide
End Sub



Private Sub imgMinimize_Click()
    Me.Hide
End Sub
Private Sub imgCheckMail_Click()
    lblCheckMail_Click
End Sub

Private Sub lblCheckMail_Click()
    CheckArrayMail
End Sub
Private Sub imgSetup_Click()
    lblSetup_click
End Sub


Private Sub lblMenu_Click(Index As Integer)



    If Index = 0 Then   'ADD
        frmAccount.Show
        frmAccount.imgAdd.Visible = True
        frmAccount.lblAdd.Visible = True
        frmAccount.lblSave.Visible = False
        frmAccount.imgSave.Visible = False
    End If
    If Index = 2 Then ' EDIT
    If lvwAccounts.ListItems.Count < 1 Then Exit Sub
        frmAccounts(lvwAccounts.SelectedItem.Index - 1).Show
        frmAccounts(lvwAccounts.SelectedItem.Index - 1).imgAdd.Visible = False
        frmAccounts(lvwAccounts.SelectedItem.Index - 1).lblAdd.Visible = False
        frmAccounts(lvwAccounts.SelectedItem.Index - 1).lblSave.Visible = True
        frmAccounts(lvwAccounts.SelectedItem.Index - 1).imgSave.Visible = True
    End If
    If Index = 1 Then
    If lvwAccounts.ListItems.Count < 1 Then Exit Sub
        MsgBox lvwAccounts.SelectedItem
        UnloadMailArray
        lvwAccounts.ListItems.Remove (lvwAccounts.SelectedItem.Index)
        RebuildMailArray
        SaveMailArray
    End If

End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim mCounter As Integer

    For mCounter = 0 To lblMenu.Count - 1
        If mCounter <> Index Then lblMenu(mCounter).ForeColor = Colors.LabelForeColor Else _
        lblMenu(mCounter).ForeColor = Colors.TitleColor
    Next mCounter

End Sub


Private Sub lblSetup_click()
    frmColors.Show
End Sub


'------------------------------
'   "CheckMail" Image Events
'------------------------------
Private Sub imgCheckMail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgCheckMail.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblCheckMail.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub imgCheckMail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCheckMail.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblCheckMail.ForeColor = Colors.ButtonForeColor
End Sub
Private Sub lblCheckMail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgCheckMail.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblCheckMail.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub lblCheckMail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCheckMail.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblCheckMail.ForeColor = Colors.ButtonForeColor
End Sub
'------------------------------
'   "Setup" Image Events
'------------------------------
Private Sub imgSetup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgSetup.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblSetup.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub imgSetup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgSetup.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblSetup.ForeColor = Colors.ButtonForeColor
End Sub
Private Sub lblSetup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgSetup.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblSetup.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub lblSetup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgSetup.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblSetup.ForeColor = Colors.ButtonForeColor
End Sub
'------------------------------
'   "Minimize" Image Events
'------------------------------
Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgMinimize.Picture = frmSkinTray.imgSkins(MinimizeDNID).Picture
End If
End Sub
Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgMinimize.Picture = frmSkinTray.imgSkins(MinimizeUPID).Picture
End If
End Sub
'------------------------------
'   "Close" Image Events
'------------------------------
Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgClose.Picture = frmSkinTray.imgSkins(CloseDNID).Picture
End If
End Sub
Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgClose.Picture = frmSkinTray.imgSkins(CloseUPID).Picture
End If
End Sub


'##################################
'   Misc Eye Candy (Focus)
'##################################
Private Sub txtdelay_LostFocus()
    txtDelay.BackColor = Colors.TextBoxBackColor
    txtDelay.ForeColor = Colors.TextBoxForeColor
End Sub
Private Sub txtdelay_GotFocus()
    txtDelay.BackColor = &H80000018
    txtDelay.ForeColor = &H80000017
End Sub
