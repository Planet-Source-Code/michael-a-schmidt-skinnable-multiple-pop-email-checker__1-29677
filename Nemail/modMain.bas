Attribute VB_Name = "modMain"
Option Explicit

' ////// Mail Setup //////////'
Public frmAccounts() As Form        ' Array of MailBoxes (Forms). Starts at 0.
Public NumAccounts As Integer       ' Number of Mailboxes. Starts at 1.
Public MailTray As New clsTray      ' Mail Tray
Public NewMail As Boolean           ' NewMail value determines system tray icon.
Public CheckIn As Integer           ' Holds "CheckIn" values for each account.

' ////// Misc Setup //////////'
Public itmX As ListItem
Public CryptO As New BlowFish


'===============================
'       -- Main --
'===============================

'===============================
'   Check Mail (Array)
'===============================
Public Sub CheckArrayMail()
Dim CountBoxes As Integer

    ' Counter. Assume at least one account exists.
    CountBoxes = 1
    If CountBoxes > frmMain.lvwAccounts.ListItems.Count Then Exit Sub

    ' Disable menu access while we are checking
    ' our mail.
    frmMain.lblMenu(0).Enabled = False
    frmMain.lblMenu(1).Enabled = False
    frmMain.lblMenu(2).Enabled = False


    ' Start checking mail, set CHECKIN to zero. No
    ' one has or can check in till we start the process.
    CheckIn = 0             ' No CheckIn
    NewMail = False         ' Assume No New Mail
    ' Change system tray icon to "Checking Mail..."
    MailTray.ChangeIcon frmSkinTray, frmSkinTray.picMail(1)
    frmToolTip.lblStats.Caption = "Checking Mail..."

    ' Loop through each account and tell each to check for mail.
    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count
        frmAccounts(CountBoxes - 1).MyBox.CheckNewMail
        frmMain.lvwAccounts.ListItems(CountBoxes).SmallIcon = 3
        CountBoxes = CountBoxes + 1
    Wend

End Sub


Public Sub UnloadMailArray()
Dim CountBoxes As Integer

    ' Destroy our array and each form referenced.
    CountBoxes = 1

    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count
        Unload frmAccounts(CountBoxes - 1)
        CountBoxes = CountBoxes + 1
        DoEvents
    Wend
    ReDim frmAccounts(0)

End Sub


Public Sub UpdateTray()
Dim CountBoxes As Integer

    ' When this sub is called by an account after checking for mail,
    ' it's "Checking In", when all accounts have checked in, update the mail icon
    ' and set appropriate tooltips.

    ' Increment CheckIn, if all accounts have checked in, continue.
    CheckIn = CheckIn + 1
    If CheckIn <> NumAccounts Then Exit Sub

    ' Set Tray Icon to New Mail or No Mail.
    If NewMail Then MailTray.ChangeIcon frmSkinTray, frmSkinTray.picMail.Item(2) _
    Else MailTray.ChangeIcon frmSkinTray, frmSkinTray.picMail.Item(0)

    CountBoxes = 1
    ' Loop through each account and view number of new messages,
    ' status, build tooltip.
    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count
    Dim MailStat, MailStats As String   ' Holds ToolTip Stuff
    
        frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(3) = frmAccounts(CountBoxes - 1).txtNewMail

        ' Set Tooltip
        MailStat = frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(3) & " Messages on " & _
                       frmMain.lvwAccounts.ListItems(CountBoxes)
        ' Set Appropriate Icons in Main ListView.
        ' Check for any errors. Set tooltip display if errors.
        If frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(3) = "[E]" Then
           MailStat = "[E] " & frmMain.lvwAccounts.ListItems(CountBoxes) & " " & frmAccounts(CountBoxes - 1).txtStatus
           frmMain.lvwAccounts.ListItems(CountBoxes).SmallIcon = 5
        ElseIf frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(3) = "[0]" Then
           frmMain.lvwAccounts.ListItems(CountBoxes).SmallIcon = 1
        Else
            frmMain.lvwAccounts.ListItems(CountBoxes).SmallIcon = 6
        End If

        ' Build ToolTip
        MailStats = MailStat & vbCrLf & MailStats
        CountBoxes = CountBoxes + 1
    Wend

    frmToolTip.lblStats = MailStats
    
    frmMain.lblMenu(0).Enabled = True
    frmMain.lblMenu(1).Enabled = True
    frmMain.lblMenu(2).Enabled = True

End Sub


Public Sub NewMailEntry(popUsername, popPassword, popServer As String)
' This only present the form, doesn't save the settings.
' frmAccount passes data, we create a new array for saving it.

    Set itmX = frmMain.lvwAccounts.ListItems.Add(, , popServer, 1, 1)
        itmX.SubItems(1) = popUsername
        itmX.SubItems(2) = CryptO.EncryptString(popPassword)

    RebuildMailArray
    SaveMailArray

End Sub


Public Sub UpdateMailList()
Dim CountBoxes As Integer

    CountBoxes = 1

    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count

        frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(1) = frmAccounts(CountBoxes - 1).txtUsername
        frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(2) = CryptO.EncryptString(frmAccounts(CountBoxes - 1).txtPassword)
        frmMain.lvwAccounts.ListItems(CountBoxes).Text = frmAccounts(CountBoxes - 1).txtServer

        CountBoxes = CountBoxes + 1
    Wend

End Sub


Public Sub SaveMailArray()
Dim CountBoxes As Integer

    CountBoxes = 1
    
    Call SaveSetting(App.Title, "SETTINGS", "ACCOUNTS", frmMain.lvwAccounts.ListItems.Count)
    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count

        Call SaveSetting(App.Title, CountBoxes, "USERNAME", frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(1))
        Call SaveSetting(App.Title, CountBoxes, "PASSWORD", frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(2))
        Call SaveSetting(App.Title, CountBoxes, "SERVER", frmMain.lvwAccounts.ListItems(CountBoxes))

        CountBoxes = CountBoxes + 1
    Wend

End Sub


Public Sub LoadMailArray()
Dim CountBoxes As Integer

    ' Load number of accounts. Starts at 1 for us.
    ' Note our array starts at 0 however.
    NumAccounts = GetSetting(App.Title, "SETTINGS", "ACCOUNTS", 0)
    CountBoxes = 1

    ' Loop through each box and build array.
    While CountBoxes <= NumAccounts

        Set itmX = frmMain.lvwAccounts.ListItems.Add(, , GetSetting(App.Title, CountBoxes, "SERVER"), 1, 1)
            itmX.SubItems(1) = GetSetting(App.Title, CountBoxes, "USERNAME")
            itmX.SubItems(2) = GetSetting(App.Title, CountBoxes, "PASSWORD")
        
        CountBoxes = CountBoxes + 1
    Wend

    RebuildMailArray

End Sub


Public Sub RebuildMailArray()
Dim CountBoxes As Integer

    CountBoxes = 1
    
    While CountBoxes <= frmMain.lvwAccounts.ListItems.Count
        
        ' Build Array of POP Boxes (Forms) based on contents of list view.
        ReDim Preserve frmAccounts(CountBoxes - 1)
        Set frmAccounts(CountBoxes - 1) = New frmAccount
        Load frmAccounts(CountBoxes - 1)

        ' Write Settings to POP Boxes (Forms)
        frmAccounts(CountBoxes - 1).txtUsername = frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(1)
        frmAccounts(CountBoxes - 1).txtPassword = CryptO.DecryptString(frmMain.lvwAccounts.ListItems(CountBoxes).SubItems(2))
        frmAccounts(CountBoxes - 1).txtServer = frmMain.lvwAccounts.ListItems(CountBoxes)

        ' Set Mailbox (Form)
        frmAccounts(CountBoxes - 1).MyBox.Password = frmAccounts(CountBoxes - 1).txtPassword
        frmAccounts(CountBoxes - 1).MyBox.User = frmAccounts(CountBoxes - 1).txtUsername
        frmAccounts(CountBoxes - 1).MyBox.Server = frmAccounts(CountBoxes - 1).txtServer
        
        Debug.Print "Array Built: " & CountBoxes & ":" & frmAccounts(CountBoxes - 1).txtServer
        CountBoxes = CountBoxes + 1
    
    Wend

    NumAccounts = frmMain.lvwAccounts.ListItems.Count
    
End Sub


