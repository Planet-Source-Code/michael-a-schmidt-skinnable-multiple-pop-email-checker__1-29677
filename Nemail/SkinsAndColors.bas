Attribute VB_Name = "SkinsAndColors"
Option Explicit

'Global Type Arrays...
Global Skins As tSkins
Global Colors As tColors

'Colors...
Type tColors
    LabelForeColor As Long
    TitleColor As Long
    ButtonForeColor As Long
    TextBoxBackColor As Long
    TextBoxForeColor As Long
    ListBoxBackColor As Long
    ButtonDownForeColor As Long
    UpdateColors As Boolean
End Type

'Skins...
Type tSkins
    UseSkins As Boolean
    SkinScheme As String
    PreviousSkin As String
    UpdateSkins As Boolean
    UseINI As Boolean
End Type

'Skin Constants...
Global Const ButtonUPID As Byte = 0
Global Const ButtonDNID As Byte = 1
Global Const CheckONID As Byte = 2
Global Const CheckOFFID As Byte = 3
Global Const CloseUPID As Byte = 4
Global Const CloseDNID As Byte = 5
Global Const MinimizeUPID As Byte = 6
Global Const MinimizeDNID As Byte = 7
Global Const RadioONID As Byte = 8
Global Const RadioOFFID As Byte = 9
Global Const RestoreUPID As Byte = 10
Global Const RestoreDNID As Byte = 11
Global Const SmallForm As Byte = 12
Global Const LargeForm As Byte = 13

Sub LoadColors()

'Loads user specific colors for the application...

On Local Error Resume Next
If Skins.UseSkins = False Then Exit Sub
Colors.LabelForeColor = Val(ReadINI("Colors", "LabelForeColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.TitleColor = Val(ReadINI("Colors", "TitleColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.ButtonForeColor = Val(ReadINI("Colors", "ButtonForeColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.TextBoxBackColor = Val(ReadINI("Colors", "TextBoxBackColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.TextBoxForeColor = Val(ReadINI("Colors", "TextBoxForeColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.ListBoxBackColor = Val(ReadINI("Colors", "ListBoxBackColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
Colors.ButtonDownForeColor = Val(ReadINI("Colors", "ButtonDownForeColor", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))

End Sub
Sub LoadSkins(Who As Form, Optional bUseINISettings As Boolean)

On Local Error GoTo LoadSkinsError

Dim x As Byte
Dim iErrorCount As Byte
Dim iResumed As Boolean

'Use or Don't Use Skins...
'If Skins.UseSkins = False Then Exit Sub

'Make sure the skin folder exists first. If it doesn't, then default to the default skin...
PreviousSkin:
If Dir$(App.Path & "\Skins\" & Skins.SkinScheme, vbDirectory) = "" Or Skins.SkinScheme = "" Then
    GoTo AllOtherForms
End If


If Who.Tag = "SKIPSKIN" Then Exit Sub
'Load skin scheme for the main menu...
If Who.Name = "frmSkinTray" Then
    Who.Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\800x600.Jpg")
    Who.imgSkins(0).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\ButtonUP.Jpg")
    Who.imgSkins(1).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\ButtonDN.Jpg")
    Who.imgSkins(2).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\CheckON.Jpg")
    Who.imgSkins(3).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\CheckOFF.Jpg")
    Who.imgSkins(4).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\CloseUP.Jpg")
    Who.imgSkins(5).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\CloseDN.Jpg")
    Who.imgSkins(6).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\MinimizeUP.Jpg")
    Who.imgSkins(7).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\MinimizeDN.Jpg")
    Who.imgSkins(8).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\RadioON.Jpg")
    Who.imgSkins(9).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\RadioOFF.Jpg")
    Who.imgSkins(10).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\RestoreUP.Jpg")
    Who.imgSkins(11).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\RestoreDN.Jpg")
    Who.imgSkins(12).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\306x184.jpg")
    Who.imgSkins(13).Picture = LoadPicture(App.Path & "\Skins\" & Skins.SkinScheme & "\500x300.jpg")
End If

'All Other Forms...
AllOtherForms:
'Debug.Print "Skinning " & Who.Name
If Who.Name <> "frmSkinTray" Then
    If Who.Tag2 = "306x184" Then Who.Picture = frmSkinTray.imgSkins(SmallForm).Picture
    If Who.Tag2 = "500x300" Then Who.Picture = frmSkinTray.imgSkins(LargeForm).Picture
    For x = 0 To Who.Controls.Count - 1
        If TypeOf Who.Controls(x) Is Image Then
            'Buttons...
            If Who.Controls(x).Tag = "Button" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(ButtonUPID).Picture
            'Minimize...
            ElseIf Who.Controls(x).Tag = "Minimize" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(MinimizeUPID).Picture
            'Restore / Maximize...
            ElseIf Who.Controls(x).Tag = "Maximize" Or Who.Controls(x).Tag = "Restore" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(RestoreUPID).Picture
            'Close...
            ElseIf Who.Controls(x).Tag = "Close" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(CloseUPID).Picture
            'CheckBox...
            ElseIf Who.Controls(x).Tag = "Check" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(CheckONID).Picture
            'Radio...
            ElseIf Who.Controls(x).Tag = "Radio" Then
                Who.Controls(x).Picture = frmSkinTray.imgSkins(RadioONID).Picture
            End If
        End If
    Next x

    'Load the skin ini settings...
    If Skins.UseINI = False Then Exit Sub
    If bUseINISettings Then
        'Caption...
        If Val(ReadINI(Who.Tag2, "CaptionLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini")) <> 0 Then
            For x = 0 To Who.Controls.Count - 1
                If Who.Controls(x).Name = "lblCaption" Then
                    Who.lblCaption.Left = Val(ReadINI(Who.Tag2, "CaptionLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                    Who.lblCaption.Top = Val(ReadINI(Who.Tag2, "CaptionTop", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                End If
            Next x
        End If
        'Close...
        If Val(Val(ReadINI(Who.Tag2, "CloseLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))) <> 0 Then
            For x = 0 To Who.Controls.Count - 1
                If Who.Controls(x).Name = "imgClose" Then
                    Who.imgClose.Left = Val(ReadINI(Who.Tag2, "CloseLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                    Who.imgClose.Top = Val(ReadINI(Who.Tag2, "CloseTop", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                End If
            Next x
        End If
        'Restore...
        If Val(Val(ReadINI(Who.Tag2, "RestoreLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))) <> 0 Then
            For x = 0 To Who.Controls.Count - 1
                If Who.Controls(x).Name = "imgRestore" Then
                    Who.imgRestore.Left = Val(ReadINI(Who.Tag2, "RestoreLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                    Who.imgRestore.Top = Val(ReadINI(Who.Tag2, "RestoreTop", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                End If
            Next x
        End If
        'Minimize...
        If Val(Val(ReadINI(Who.Tag2, "MinimizeLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))) <> 0 Then
            For x = 0 To Who.Controls.Count - 1
                If Who.Controls(x).Name = "imgMinimize" Then
                    Who.imgMinimize.Left = Val(ReadINI(Who.Tag2, "MinimizeLeft", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                    Who.imgMinimize.Top = Val(ReadINI(Who.Tag2, "MinimizeTop", App.Path & "\Skins\" & Skins.SkinScheme & "\" & Skins.SkinScheme & ".Ini"))
                End If
            Next x
        End If
    End If
End If

'Re-Load the color settings for this skin...
If iResumed Then
    Colors.UpdateColors = True
End If

Exit Sub



LoadSkinsError:
    iErrorCount = iErrorCount + 1
    If iErrorCount > 1 Then
        'Previous skin didn't work either so just write the error and proceed...
        Call WriteToErrorLog("SkinsAndColors", "LoadSkinsError", Err.Description, Err.Number, False)
        Exit Sub
    Else
        'Try using the previous skin...
        Call WriteToErrorLog("SkinsAndColors", "LoadSkinsError", Err.Description, Err.Number, False, "Unable to load this skin. Reverting back to the previous skin.", vbInformation)
        Skins.SkinScheme = Skins.PreviousSkin
        iResumed = True
        Resume PreviousSkin
    End If

End Sub
Sub SetColors(Who As Form)

On Local Error Resume Next

Dim x As Integer

If Skins.UseINI = False Then Exit Sub
'Set label and button label fore colors...
For x = 0 To Who.Controls.Count - 1
    If InStr(LCase$(Who.Controls(x).Tag), "nocolorchange") = 0 Then
        'Label Colors...
        If Who.Controls(x).Tag = "Label" Then
            Who.Controls(x).ForeColor = Colors.LabelForeColor
        'Title Color...
        ElseIf Who.Controls(x).Tag = "TitleColor" Then
            Who.Controls(x).ForeColor = Colors.TitleColor
        'Button Label Colors...
        ElseIf Who.Controls(x).Tag = "ButtonLabel" Then
            Who.Controls(x).ForeColor = Colors.ButtonForeColor
        'Textbox ForeGround and BackGround Colors...
        ElseIf TypeOf Who.Controls(x) Is TextBox Then
            Who.Controls(x).ForeColor = Colors.TextBoxForeColor
            Who.Controls(x).BackColor = Colors.TextBoxBackColor
        'List and combo box BackGround Colors...
        ElseIf TypeOf Who.Controls(x) Is ListBox Or TypeOf Who.Controls(x) Is ComboBox Or TypeOf Who.Controls(x) Is ListView Then
            Who.Controls(x).BackColor = Colors.ListBoxBackColor
            Who.Controls(x).ForeColor = Colors.TextBoxForeColor
        End If
    End If
Next x

'Check for any errors and write them to the ErrorLog.Txt file...
If Err.Number > 0 Then
    Call WriteToErrorLog("SkinsAndColors", "SetColors", Err.Description, Err.Number, False)
End If

End Sub
