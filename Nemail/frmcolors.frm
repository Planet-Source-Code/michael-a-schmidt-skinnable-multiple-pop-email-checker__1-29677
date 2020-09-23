VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Color Settings"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   7980
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picColorPalette 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   540
      MouseIcon       =   "frmcolors.frx":0000
      Picture         =   "frmcolors.frx":0152
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   2
      Top             =   690
      Width           =   3015
      Begin VB.PictureBox picColorsSquare 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1605
         Left            =   420
         MouseIcon       =   "frmcolors.frx":4822
         Picture         =   "frmcolors.frx":4974
         ScaleHeight     =   103
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8010
      Top             =   180
   End
   Begin VB.PictureBox picControls 
      Height          =   3015
      Left            =   3840
      ScaleHeight     =   2955
      ScaleWidth      =   2985
      TabIndex        =   6
      Top             =   690
      Width           =   3045
      Begin VB.ComboBox cboSkinSchemes 
         Height          =   315
         ItemData        =   "frmcolors.frx":83F1
         Left            =   540
         List            =   "frmcolors.frx":83F3
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2190
         Width           =   2295
      End
      Begin VB.ComboBox lstListBoxBackGroundColor 
         Height          =   315
         ItemData        =   "frmcolors.frx":83F5
         Left            =   540
         List            =   "frmcolors.frx":83FC
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Listbox Background Color"
         ToolTipText     =   "Click here to set the background color for all list boxes"
         Top             =   1815
         Width           =   2295
      End
      Begin VB.TextBox txtTextBoxForeGroundColor 
         Height          =   285
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Textbox Foreground Color"
         ToolTipText     =   "Click here to set the text color for all text boxes"
         Top             =   1455
         Width           =   2295
      End
      Begin VB.TextBox txtTextBoxBackGroundColor 
         Height          =   285
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Textbox Background Color"
         ToolTipText     =   "Click here to set the background color for all text boxes"
         Top             =   1095
         Width           =   2295
      End
      Begin VB.Label lblHideSolidColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Solid Colors"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Tag             =   "Label"
         ToolTipText     =   "Click here to hide and show the solid colors palette"
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label lblAutoApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Apply"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2070
         TabIndex        =   15
         Tag             =   "Label"
         ToolTipText     =   "Click here for all color changes to take effect immediately"
         Top             =   2640
         Width           =   765
      End
      Begin VB.Image imgShowHideSolidColors 
         Height          =   225
         Left            =   150
         ToolTipText     =   "Click here to show and hide the solid colors"
         Top             =   2640
         Width           =   225
      End
      Begin VB.Image imgAutoApply 
         Height          =   225
         Left            =   1800
         ToolTipText     =   "Click here for all color changes to take effect immediately"
         Top             =   2640
         Width           =   225
      End
      Begin VB.Image imgSelected 
         Height          =   210
         Index           =   6
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   420
         Width           =   210
      End
      Begin VB.Label lblTitleColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   14
         Tag             =   "TitleColor"
         ToolTipText     =   "Click here to set the color of the window name or title"
         Top             =   420
         Width           =   705
      End
      Begin VB.Image imgSelected 
         Height          =   225
         Index           =   5
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   2190
         Width           =   225
      End
      Begin VB.Image imgSelected 
         Height          =   225
         Index           =   4
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   1830
         Width           =   225
      End
      Begin VB.Image imgSelected 
         Height          =   225
         Index           =   3
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   1470
         Width           =   225
      End
      Begin VB.Image imgSelected 
         Height          =   225
         Index           =   2
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   1110
         Width           =   225
      End
      Begin VB.Image imgSelected 
         Height          =   225
         Index           =   1
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   750
         Width           =   225
      End
      Begin VB.Label lblLabelColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Color"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Tag             =   "Label"
         ToolTipText     =   "Click here to set the color of all labels"
         Top             =   90
         Width           =   795
      End
      Begin VB.Label lblButtonLabelColor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Button Text Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   630
         TabIndex        =   10
         Tag             =   "ButtonLabel"
         ToolTipText     =   "Click here to set the color of all text on buttons"
         Top             =   750
         Width           =   1245
      End
      Begin VB.Image imgSelected 
         Height          =   210
         Index           =   0
         Left            =   150
         ToolTipText     =   "Click on these radio buttons to select what's to the right of them"
         Top             =   90
         Width           =   210
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   540
         Stretch         =   -1  'True
         Tag             =   "Button"
         ToolTipText     =   "Click here to set the color of all text on buttons"
         Top             =   705
         Width           =   1425
      End
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3690
      TabIndex        =   12
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to apply your changes and close this window"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Left            =   6930
      Tag             =   "Minimize"
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   7200
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   285
   End
   Begin VB.Label lblWindowsColors 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More Colors..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1890
      TabIndex        =   5
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to go to the Windows Colors screen"
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label lblApply 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6480
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to apply your color changes and leave the window open"
      Top             =   3960
      Width           =   405
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Settings"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Tag             =   "TitleColor"
      Top             =   330
      Width           =   1005
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4980
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Click here to cancel your changes or to close this window"
      Top             =   3960
      Width           =   525
   End
   Begin VB.Image imgExit 
      Height          =   345
      Left            =   4530
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to cancel your changes or to close this window"
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Image imgApply 
      Enabled         =   0   'False
      Height          =   345
      Left            =   5940
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to apply your color changes and leave the window open"
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Image imgWindowsColors 
      Height          =   345
      Left            =   1710
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to go to the Windows Colors screen"
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   345
      Left            =   3120
      Stretch         =   -1  'True
      Tag             =   "Button"
      ToolTipText     =   "Click here to apply your changes and close this window"
      Top             =   3900
      Width           =   1395
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Dim sControlSelected As String
Private FormIsLoading As Boolean
Private m_Tag2 As String
Sub SaveChanges()

On Local Error Resume Next

'Save the color settings to the skin scheme ini file...
Call WriteINI("Colors", "LabelForeColor", lblLabelColor.ForeColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")
Call WriteINI("Colors", "TitleColor", lblTitleColor.ForeColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")
Call WriteINI("Colors", "ButtonForeColor", lblButtonLabelColor.ForeColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")
Call WriteINI("Colors", "TextBoxBackColor", txtTextBoxBackGroundColor.BackColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")
Call WriteINI("Colors", "TextBoxForeColor", txtTextBoxForeGroundColor.ForeColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")
Call WriteINI("Colors", "ListBoxBackColor", lstListBoxBackGroundColor.BackColor, App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")

iDirty = False
Colors.UpdateColors = True
lblExit.Caption = "Close"

End Sub
Private Sub cboSkinSchemes_Click()

On Local Error Resume Next

'Exit if form is loading...
If FormIsLoading Then Exit Sub

'Remember the current skin scheme in case we have to revert back to it because of errors...
Skins.PreviousSkin = Skins.SkinScheme

'Apply The New Scheme...
Skins.SkinScheme = cboSkinSchemes
Skins.UseSkins = True

'Save Settings...
Call WriteINI("Skins", "UseSkins", Trim$(Skins.UseSkins), QuickRef.UserINIFileName)
Call WriteINI("Skins", "SkinScheme", Skins.SkinScheme, QuickRef.UserINIFileName)

'Tells all other open forms to update themselves...
Skins.UpdateSkins = True
Colors.UpdateColors = True

End Sub
Private Sub cboSkinSchemes_GotFocus()

Call imgSelected_Click(5)

End Sub

Private Sub cboSkinSchemes_KeyPress(KeyAscii As Integer)

'Disable keystrokes...
KeyAscii = 0

End Sub
Private Sub Form_Load()

On Local Error Resume Next

'Prevents recursion on the Skins ComboBox...
FormIsLoading = True

Call LoadINISettings
Call LoadSkins(Me, True)
Call SetColors(Me)

'Load All Available Skin Schemes...
If LoadAllSkinSchemes = False Then
    cboSkinSchemes.Enabled = False
End If

'Start out with the label being selected...
sControlSelected = "Label"

'Form Coordinates...
Me.Width = 7500
Me.Height = 4505

iDirty = False
FormIsLoading = False

End Sub
Public Property Get Tag2() As String

Tag2 = m_Tag2

End Property
Public Property Let Tag2(sValue As String)

m_Tag2 = sValue

End Property
Function LoadAllSkinSchemes() As Boolean

'This routine loads all skin schemes into the ComboBox on this form so that the user can select one. The Function returns
'False if there was a problem loading the skins from the .INI file or any of the Skins subdirectories doesn't exist. If
'a Skin subdirectory doesn't exist, it isn't added to the ComboBox for the user to select. If at least one or more Skin
'subdirectories DO exist, the function will return true and give the user the ability to select from those skins...

On Local Error GoTo LoadAllSkinSchemesError

Dim x As Byte
Dim sInput As String
Dim FileFree As Byte

'Load All Skin Schemes Found In The Skins.INI File... <-- OLD!!!
FileFree = FreeFile
cboSkinSchemes.Clear
Dim iName As String   ' Directory / File
Dim Found As Boolean
Dim SkinPath As String
    x = 0
    ' if directory doesn't exist create it. <<TODO>>
    SkinPath = App.Path & "\Skins\"
    iName = Dir(SkinPath, vbDirectory + vbHidden + vbArchive + vbSystem + vbReadOnly)  ' Retrieve the first entry.

    Do While iName <> ""              ' Start the loop.

    If iName <> "." And iName <> ".." And iName <> "System Volume Information" Then
        If (GetAttr(SkinPath & iName) And vbDirectory) = vbDirectory Then
            cboSkinSchemes.AddItem iName
            x = x + 1

        End If
    End If

    iName = Dir                       ' Get next entry.
    Loop



'Set The Skin Scheme ComboBox To The Current Scheme...
If Skins.UseSkins = True And Skins.SkinScheme <> "" Then
    For x = 0 To cboSkinSchemes.ListCount - 1
        If LCase$(cboSkinSchemes.List(x)) = LCase$(Skins.SkinScheme) Then
            cboSkinSchemes.ListIndex = x
            Exit For
        End If
    Next x
End If

LoadAllSkinSchemes = True
Exit Function



LoadAllSkinSchemesError:
    Call WriteToErrorLog(Me.Name, "LoadAllSkinSchemesError", Err.Description, Err.Number, False)
    Exit Function

End Function
Sub LoadINISettings()

'Set the custom Tag2 property (for skins)...
Me.Tag2 = "500x300"

'Controls PictureBox...
picControls.Picture = frmSkinTray.Picture

'Color Palettes Pictures...
If Dir$(App.Path & "\Bitmaps\ColorSquare.Jpg") <> "" Then
    picColorsSquare.Picture = LoadPicture(App.Path & "\Bitmaps\ColorSquare.Jpg")
End If
If Dir$(App.Path & "\Bitmaps\ColorPicker.Jpg") <> "" Then
    picColorPalette.Picture = LoadPicture(App.Path & "\Bitmaps\ColorPicker.Jpg")
End If

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left", QuickRef.UserINIFileName))
Me.Top = Val(ReadINI(Me.Name, "Top", QuickRef.UserINIFileName))

'Auto Apply...
If ReadINI(Me.Name, "AutoApply", QuickRef.UserINIFileName) = True Then
    imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture
Else
    imgAutoApply.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture
End If

'Show / Hide Solid Colors...
imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture

'Default the radio buttons on this form...
Call imgSelected_Click(0)

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

Dim x As Byte

'Prompt to save first...
If iDirty Then
    x = MsgBox("Save changes before exiting?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            Call SaveChanges
        Case vbCancel
            Cancel = True
            QuickRef.CancelOperation = True
            Exit Sub
    End Select
End If

'Save INI Settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left, QuickRef.UserINIFileName)
Call WriteINI(Me.Name, "Top", Me.Top, QuickRef.UserINIFileName)

'Auto Apply...
Call WriteINI(Me.Name, "AutoApply", (imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture), QuickRef.UserINIFileName)

End Sub
Private Sub Image1_Click()

'Remember the control selected...
sControlSelected = "ButtonLabel"

'Update the radio button...
Call imgSelected_Click(1)

End Sub

Private Sub imgAutoApply_Click()

lblAutoApply_Click

End Sub
Private Sub imgApply_Click()

Call SaveChanges

End Sub

Private Sub imgApply_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgApply.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblApply.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub imgApply_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgApply.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblApply.ForeColor = Colors.ButtonForeColor

End Sub

Private Sub imgClose_Click()

Unload Me

End Sub

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
Private Sub imgMinimize_Click()
Me.Hide
End Sub

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
Private Sub imgSelected_Click(Index As Integer)

Dim x As Byte

'Clear the radio buttons...
For x = 0 To 6
    imgSelected(x).Picture = frmSkinTray.imgSkins(RadioOFFID).Picture
Next x

'Update the radio buttons...
imgSelected(Index).Picture = frmSkinTray.imgSkins(RadioONID).Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected = "Label"
    Case 1
        sControlSelected = "ButtonLabel"
    Case 2
        sControlSelected = "TextboxBackgroundColor"
    Case 3
        sControlSelected = "TextboxForegroundColor"
    Case 4
        sControlSelected = "ListboxBackgroundColor"
    Case 5
        sControlSelected = "SkinSchemes"
    Case 6
        sControlSelected = "TitleColor"
End Select

End Sub
Private Sub imgShowHideSolidColors_Click()

lblHideSolidColors_Click

End Sub

Private Sub imgWindowsColors_Click()

lblWindowsColors_Click

End Sub
Private Sub imgWindowsColors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgWindowsColors.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblWindowsColors.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub imgWindowsColors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgWindowsColors.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblWindowsColors.ForeColor = Colors.ButtonForeColor

End Sub
Private Sub lblAutoApply_Click()

'Show / Hide Solid Colors...
If imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture Then
    imgAutoApply.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture
Else
    imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture
    Call lblApply_Click
End If

End Sub
Private Sub lblButtonLabelColor_Click()

'Remember the control selected...
sControlSelected = "ButtonLabel"

'Update the radio button...
Call imgSelected_Click(1)

End Sub
Private Sub lblLabelColor_Click()

'Remember the control selected...
sControlSelected = "Label"

'Update the radio button...
Call imgSelected_Click(0)

End Sub

Private Sub lblTitleColor_Click()

'Remember the control selected...
sControlSelected = "TitleColor"

'Update the radio button...
Call imgSelected_Click(6)

End Sub
Private Sub lblWindowsColors_Click()

On Local Error Resume Next

'Set the frmSkinTray.Dialog boxes color to the color of the control that is selected...
If sControlSelected = "Label" Then
    frmSkinTray.Dialog.Color = lblLabelColor.ForeColor
ElseIf sControlSelected = "ButtonLabel" Then
    frmSkinTray.Dialog.Color = lblButtonLabelColor.ForeColor
ElseIf sControlSelected = "TextboxBackgroundColor" Then
    frmSkinTray.Dialog.Color = txtTextBoxBackGroundColor.BackColor
ElseIf sControlSelected = "TextboxForegroundColor" Then
    frmSkinTray.Dialog.Color = txtTextBoxForeGroundColor.ForeColor
ElseIf sControlSelected = "TitleColor" Then
    frmSkinTray.Dialog.Color = lblTitleColor.ForeColor
End If

'Show the color frmSkinTray.Dialog box...
frmSkinTray.Dialog.Flags = cdlCCFullOpen Or cdlCCRGBInit
frmSkinTray.Dialog.ShowColor
If Err.Number > 0 Then Exit Sub

'Set the color to the control currently selected...
If sControlSelected = "Label" Then
    lblLabelColor.ForeColor = frmSkinTray.Dialog.Color
ElseIf sControlSelected = "ButtonLabel" Then
    lblButtonLabelColor.ForeColor = frmSkinTray.Dialog.Color
ElseIf sControlSelected = "TextboxBackgroundColor" Then
    txtTextBoxBackGroundColor.BackColor = frmSkinTray.Dialog.Color
    txtTextBoxForeGroundColor.BackColor = frmSkinTray.Dialog.Color
ElseIf sControlSelected = "TextboxForegroundColor" Then
    txtTextBoxForeGroundColor.ForeColor = frmSkinTray.Dialog.Color
    txtTextBoxBackGroundColor.ForeColor = frmSkinTray.Dialog.Color
ElseIf sControlSelected = "ListboxBackgroundColor" Then
    lstListBoxBackGroundColor.BackColor = frmSkinTray.Dialog.Color
    cboSkinSchemes.BackColor = lstListBoxBackGroundColor.BackColor
ElseIf sControlSelected = "TitleColor" Then
    lblTitleColor.ForeColor = frmSkinTray.Dialog.Color
End If

'Auto Apply...
If imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture Then
    Call lblApply_Click
    Exit Sub
End If

iDirty = True

End Sub
Private Sub lblWindowsColors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgWindowsColors.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblWindowsColors.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub lblApply_Click()

Call SaveChanges

End Sub
Private Sub lblApply_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgApply.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblApply.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub lblApply_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgApply.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblApply.ForeColor = Colors.ButtonForeColor

End Sub
Private Sub lblHideSolidColors_Click()

'Show / Hide Solid Colors...
If imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckONID).Picture Then
    imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckOFFID).Picture
    lblHideSolidColors.Caption = "Show Solid Colors"
Else
    imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckONID).Picture
    lblHideSolidColors.Caption = "Hide Solid Colors"
End If
picColorsSquare.Visible = (imgShowHideSolidColors.Picture = frmSkinTray.imgSkins(CheckONID).Picture)

End Sub
Private Sub lblWindowsColors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgWindowsColors.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblWindowsColors.ForeColor = Colors.ButtonForeColor

End Sub
Private Sub lstListBoxBackGroundColor_GotFocus()

'Remember the control selected...
sControlSelected = "ListboxBackgroundColor"

'Update the radio button...
Call imgSelected_Click(4)

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblExit.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgExit.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblExit.ForeColor = Colors.ButtonForeColor

End Sub
Private Sub imgSave_Click()

lblsave_Click

End Sub

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
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub lblExit_Click()

Unload Me

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = frmSkinTray.imgSkins(ButtonDNID).Picture
    lblExit.ForeColor = QBColor(Colors.ButtonDownForeColor)
End If

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

imgExit.Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
lblExit.ForeColor = Colors.ButtonForeColor

End Sub
Private Sub lblsave_Click()

Call SaveChanges

Unload Me

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
Private Sub picColorPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Local Error Resume Next

'Change the label color...
If sControlSelected = "Label" Then
    lblLabelColor.ForeColor = picColorPalette.Point(x, y)
ElseIf sControlSelected = "ButtonLabel" Then
    lblButtonLabelColor.ForeColor = picColorPalette.Point(x, y)
ElseIf sControlSelected = "TextboxBackgroundColor" Then
    txtTextBoxBackGroundColor.BackColor = picColorPalette.Point(x, y)
    txtTextBoxForeGroundColor.BackColor = picColorPalette.Point(x, y)
ElseIf sControlSelected = "TextboxForegroundColor" Then
    txtTextBoxForeGroundColor.ForeColor = picColorPalette.Point(x, y)
    txtTextBoxBackGroundColor.ForeColor = picColorPalette.Point(x, y)
    lstListBoxBackGroundColor.ForeColor = picColorPalette.Point(x, y)
ElseIf sControlSelected = "ListboxBackgroundColor" Then
    lstListBoxBackGroundColor.BackColor = picColorPalette.Point(x, y)
    cboSkinSchemes.BackColor = picColorPalette.Point(x, y)
ElseIf sControlSelected = "TitleColor" Then
    lblTitleColor.ForeColor = picColorPalette.Point(x, y)
    lblCaption.ForeColor = picColorPalette.Point(x, y)
End If

iDirty = True

End Sub
Private Sub picColorPalette_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Local Error Resume Next

'Change the label color...
If Button = vbLeftButton Then
    If sControlSelected = "Label" Then
        lblLabelColor.ForeColor = picColorPalette.Point(x, y)
    ElseIf sControlSelected = "ButtonLabel" Then
        lblButtonLabelColor.ForeColor = picColorPalette.Point(x, y)
    ElseIf sControlSelected = "TextboxBackgroundColor" Then
        txtTextBoxBackGroundColor.BackColor = picColorPalette.Point(x, y)
        txtTextBoxForeGroundColor.BackColor = picColorPalette.Point(x, y)
    ElseIf sControlSelected = "TextboxForegroundColor" Then
        txtTextBoxForeGroundColor.ForeColor = picColorPalette.Point(x, y)
        txtTextBoxBackGroundColor.ForeColor = picColorPalette.Point(x, y)
        lstListBoxBackGroundColor.ForeColor = picColorPalette.Point(x, y)
    ElseIf sControlSelected = "ListboxBackgroundColor" Then
        lstListBoxBackGroundColor.BackColor = picColorPalette.Point(x, y)
        cboSkinSchemes.BackColor = picColorPalette.Point(x, y)
    ElseIf sControlSelected = "TitleColor" Then
        lblTitleColor.ForeColor = picColorPalette.Point(x, y)
        lblCaption.ForeColor = picColorPalette.Point(x, y)
    End If
    iDirty = True
    DoEvents
End If

End Sub
Private Sub picColorPalette_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'Auto Apply...
If imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture Then
    Call lblApply_Click
End If

End Sub
Private Sub picColorsSquare_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Local Error Resume Next

'Change the label color...
If Button = vbLeftButton Then
    If sControlSelected = "Label" Then
        lblLabelColor.ForeColor = picColorsSquare.Point(x, y)
    ElseIf sControlSelected = "ButtonLabel" Then
        lblButtonLabelColor.ForeColor = picColorsSquare.Point(x, y)
    ElseIf sControlSelected = "TextboxBackgroundColor" Then
        txtTextBoxBackGroundColor.BackColor = picColorsSquare.Point(x, y)
        txtTextBoxForeGroundColor.BackColor = picColorsSquare.Point(x, y)
    ElseIf sControlSelected = "TextboxForegroundColor" Then
        txtTextBoxForeGroundColor.ForeColor = picColorsSquare.Point(x, y)
        txtTextBoxBackGroundColor.ForeColor = picColorsSquare.Point(x, y)
        lstListBoxBackGroundColor.ForeColor = picColorsSquare.Point(x, y)
    ElseIf sControlSelected = "ListboxBackgroundColor" Then
        lstListBoxBackGroundColor.BackColor = picColorsSquare.Point(x, y)
        cboSkinSchemes.BackColor = picColorsSquare.Point(x, y)
    ElseIf sControlSelected = "TitleColor" Then
        lblTitleColor.ForeColor = picColorsSquare.Point(x, y)
        lblCaption.ForeColor = picColorsSquare.Point(x, y)
    End If
    iDirty = True
End If

End Sub
Private Sub picColorsSquare_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'Auto Apply...
If imgAutoApply.Picture = frmSkinTray.imgSkins(CheckONID).Picture Then
    Call lblApply_Click
End If

End Sub

Private Sub picControls_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Apply...
If imgApply.Enabled = False And iDirty = True Then
    imgApply.Enabled = True
    lblApply.Enabled = True
ElseIf imgApply.Enabled = True And iDirty = False Then
    imgApply.Enabled = False
    lblApply.Enabled = False
End If

End Sub
Private Sub txtTextBoxBackGroundColor_Click()

'Remember the control selected...
sControlSelected = "TextboxBackgroundColor"

'Update the radio button...
Call imgSelected_Click(2)

End Sub
Private Sub txtTextBoxForeGroundColor_Click()

'Remember the control selected...
sControlSelected = "TextboxForegroundColor"

'Update the radio button...
Call imgSelected_Click(3)

End Sub
