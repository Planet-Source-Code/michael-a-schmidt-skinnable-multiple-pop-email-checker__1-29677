Attribute VB_Name = "StartUp"
Option Explicit

Sub Main()
On Local Error Resume Next

    If App.PrevInstance Then End

'Local INI...
If Dir$(App.Path & "\Local.Ini") <> "" Then
    QuickRef.UserINIFileName = App.Path & "\Local.Ini"
End If

'Get Database Connection Information...
If Trim(ReadINI("Database", "DatabaseLocation", QuickRef.UserINIFileName)) <> "" Then
    QuickRef.DBFileName = ReadINI("Database", "DatabaseLocation", QuickRef.UserINIFileName) & "SkinDemo.Mdb"
Else
    QuickRef.DBFileName = App.Path & "\SkinDemo.Mdb"
End If
If Dir$(QuickRef.DBFileName) = "" Then
    'MsgBox "Unable to establish a connection to the Database or to the Server. Contact your system administrator for help.", vbCritical, "No Connection..."
    'End
End If

'Load the last skin scheme (if we are using skins, that is)...
Skins.UseSkins = ReadINI("Skins", "UseSkins", QuickRef.UserINIFileName)
If Skins.UseSkins Then
    Skins.SkinScheme = ReadINI("Skins", "SkinScheme", QuickRef.UserINIFileName)
    Skins.PreviousSkin = Skins.SkinScheme
    If Trim$(Skins.SkinScheme) = "" Or Dir$(App.Path & "\Skins\" & Skins.SkinScheme, vbDirectory) = "" Then
        'MsgBox "Unable to use the Skin " & Chr$(34) & Skins.SkinScheme & Chr$(34) & ". The skin folder doesn't exist. You can go to the " & Chr$(34) & "Colors Screen" & Chr$(34) & " and specify another " & Chr$(34) & "Color Scheme" & Chr$(34) & " and " & Chr$(34) & "Skin Scheme" & Chr$(34) & " from there.", vbInformation, "Skin Folder..."
        Skins.SkinScheme = ""
        Skins.PreviousSkin = ""
        Dim SkinPath As String
        SkinPath = App.Path & "\Skins\"
        If Dir(SkinPath) = "" Then MkDir SkinPath
        Skins.UseSkins = True '<<TODO>> write .INI file...
        Skins.UseINI = False
    Else
        Skins.UseSkins = True
        Skins.UseINI = True
    End If
End If

Call LoadColors
'Skins.UseSkins = True
'Main Menu...

    ' Key used for passwords. <<TODO>>
    ' Incorporate Username / Password / Server into Key.
    CryptO.Key = "CI3894UWRE3H32F89HHDSN34"

    Load frmMain        ' Load Main
    'Load frmAbout       ' Load About ( Set SKIN )
    LoadMailArray       ' Load Mail Settings
    
    ' Setup new account, or check mail.
    If NumAccounts = 0 Then
        frmAccount.Show
        frmAccount.lblAdd.Visible = True
        frmAccount.imgAdd.Visible = True
        frmAccount.lblSave.Visible = False
        frmAccount.imgSave.Visible = False
    Else
        CheckArrayMail
    End If

End Sub
