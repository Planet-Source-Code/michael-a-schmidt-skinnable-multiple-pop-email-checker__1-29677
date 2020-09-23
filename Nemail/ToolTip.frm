VERSION 5.00
Begin VB.Form frmToolTip 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   330
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3480
   ControlBox      =   0   'False
   FillColor       =   &H80000018&
   ForeColor       =   &H80000018&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "SKIPSKIN"
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[STATUS]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   690
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub Form_Load()
    ' Set our window position to topmost window.
    SetWindowPos hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H1
    Me.Hide
End Sub

Public Sub lblStats_Change()

    ' Resize our height based on the automatically
    ' resizing label. Note that if we have only one line
    ' of text in the label, the overall size is too small
    ' so we hardcode anything less than 360 pixels to 360 pixels.
    
    ' Width & Height
    Me.Height = lblStats.Height
    If Me.Height < 360 Then Me.Height = 360
    Me.Width = lblStats.Width + 500

    ' Window Positioning
    frmToolTip.Left = Screen.Width - frmToolTip.Width
    frmToolTip.Top = Screen.Height - frmToolTip.Height - 420


End Sub

