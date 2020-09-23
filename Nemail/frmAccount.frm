VERSION 5.00
Begin VB.Form frmAccount 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "frmAccount.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   690
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   990
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   2385
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   990
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   2385
   End
   Begin Nemail.VBMail MyBox 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2370
      TabIndex        =   4
      ToolTipText     =   "Enter your POP server here."
      Top             =   1740
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2370
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter your password here."
      Top             =   1200
      Width           =   2025
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2370
      TabIndex        =   2
      ToolTipText     =   "Enter your login name here."
      Top             =   660
      Width           =   2025
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3615
      TabIndex        =   11
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to add your account."
      Top             =   2250
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Tag             =   "Label"
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Tag             =   "Label"
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   5
      Tag             =   "Label"
      Top             =   420
      Width           =   780
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3570
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to save your settings."
      Top             =   2250
      Width           =   405
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to cancel."
      Top             =   2250
      Width           =   525
   End
   Begin VB.Image imgSave 
      Height          =   345
      Left            =   3060
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to save your settings."
      Top             =   2190
      Width           =   1365
   End
   Begin VB.Image imgCancel 
      Height          =   345
      Left            =   1680
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to cancel."
      Top             =   2190
      Width           =   1365
   End
   Begin VB.Image imgAdd 
      Height          =   345
      Left            =   3060
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to add your account."
      Top             =   2190
      Width           =   1365
   End
End
Attribute VB_Name = "frmAccount"
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

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save this form's settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

On Local Error Resume Next

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left, QuickRef.UserINIFileName)
Call WriteINI(Me.Name, "Top", Me.Top, QuickRef.UserINIFileName)

End Sub

Private Sub imgCancel_Click()

lblCancel_Click

End Sub


Private Sub lblCancel_Click()

    Call SaveINISettings
    Me.Hide

End Sub


Private Sub imgSave_Click()
    Call lblsave_Click
End Sub

Private Sub imgAdd_Click()
    Call lblAdd_Click
End Sub

Private Sub lblAdd_Click()
    ' Set Mailbox Settings
    Call NewMailEntry(txtUsername, txtPassword, txtServer)
    Unload Me
End Sub

Private Sub lblsave_Click()

    ' Set Mailbox Settings
    MyBox.User = txtUsername
    MyBox.Password = txtPassword
    MyBox.Server = txtServer
    
    ' Save Settings
    UpdateMailList
    SaveMailArray
    Me.Hide

End Sub


'##################################
'   Misc Eye Candy (Focus)
'##################################
Private Sub txtUsername_LostFocus()
    Call SetColors(Me)
End Sub
Private Sub txtUsername_GotFocus()
    txtUsername.BackColor = &H80000018
    txtUsername.ForeColor = &H80000017
End Sub
Private Sub txtPassword_LostFocus()
    Call SetColors(Me)
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.BackColor = &H80000018
    txtPassword.ForeColor = &H80000017
End Sub
Private Sub txtServer_LostFocus()
    Call SetColors(Me)
End Sub
Private Sub txtServer_GotFocus()
    txtServer.BackColor = &H80000018
    txtServer.ForeColor = &H80000017
End Sub


'================================
'   OBJECT Event New Mail
'================================
Private Sub MYBox_NewMail(NumMsgs As Integer)

    ' Set our textbox to show how many mails.
    ' This box is called later on by our modMAIN
    ' to build our tooltip of new messages.
    txtNewMail = "[" & NumMsgs & "]"

    txtStatus = NumMsgs & " New Message(s)"
    
    ' Every mail check NewMail is set to false,
    ' based on if it's true or not in the end,
    ' we set the appropriate icon in the tray.
    NewMail = True
    UpdateTray

End Sub


'================================
'   OBJECT Event Noisy
'================================
Private Sub MYBox_Noisy(POPresponse As String)

    txtLog = POPresponse & vbCrLf & txtLog

End Sub


'================================
'   OBJECT Event No Mail
'================================
Private Sub MYBox_NoMail()

    txtNewMail = "[0]"
    txtStatus = "No New Mail"
    UpdateTray

End Sub


'================================
'   OBJECT Event Error
'================================
Private Sub MYBox_SockError(ErrorStats As String)

    txtNewMail = "[E]"
    txtStatus = ErrorStats
    UpdateTray

End Sub


'------------------------------
'   "Cancel" Image Events
'------------------------------
Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgCancel.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblCancel.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCancel.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblCancel.ForeColor = Colors.ButtonForeColor
End Sub
Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgCancel.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblCancel.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCancel.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblCancel.ForeColor = Colors.ButtonForeColor
End Sub
'------------------------------
'   "Add" Image Events
'------------------------------
Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgAdd.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblAdd.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgAdd.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblAdd.ForeColor = Colors.ButtonForeColor
End Sub
Private Sub lblAdd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgAdd.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblAdd.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub lblAdd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgAdd.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblAdd.ForeColor = Colors.ButtonForeColor
End Sub
'------------------------------
'   "Edit" Image Events
'------------------------------
Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgSave.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblSave.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgSave.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblSave.ForeColor = Colors.ButtonForeColor
End Sub
Private Sub lblsave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    imgSave.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblSave.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If
End Sub
Private Sub lblsave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgSave.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblSave.ForeColor = Colors.ButtonForeColor
End Sub

