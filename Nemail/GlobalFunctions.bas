Attribute VB_Name = "GlobalFunctions"
Option Explicit

'Misc Variables...
Global DB As Database
Global RS As Recordset
Global sSQL As String

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'INI File Functions...
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Always on top...
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Public Sub AlwaysOnTop(Who As Form, iPosition As Boolean)

Dim lFlag As Long

'On top or not on top...
If iPosition Then
    lFlag = -1
Else
    lFlag = -2
End If

'Call the API to make the form on or not on top...
Call SetWindowPos(Who.hwnd, lFlag, Who.Left / Screen.TwipsPerPixelX, Who.Top / Screen.TwipsPerPixelY, Who.Width / Screen.TwipsPerPixelX, Who.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW)

End Sub
'Sub AlignFormToForm(CallingForm As Form, AligningForm As Form)

'On Local Error Resume Next

'Align the helper form to the form it is associated with...
'AligningForm.Left = CallingForm.Left + CallingForm.Width + 20
'AligningForm.Top = CallingForm.Top

'Make sure the helper form is visible...
'If AligningForm.Left >= frmSkinTray.Width - 600 Then
'    AligningForm.Left = (frmSkinTray.Width - AligningForm.Width) - 180
'    AligningForm.WindowState = vbNormal
'    AligningForm.ZOrder
'End If

'End Sub
Sub CloseAllOpenWindows()

On Local Error Resume Next

Dim x As Byte

'Close all currently open forms...
For x = 1 To Forms.Count - 1
    If Forms(x).Name <> "frmSkinTray" Then
        Unload Forms(x)
    End If
Next x

End Sub

Function GetAvailableDriveSpace(sDriveLetter As String) As String

'This function returns the amount of available drive space for any given drive...

On Local Error Resume Next

Dim sSpace As String

sSpace = FS.Drives(sDriveLetter).AvailableSpace

If Err.Number > 0 Then
    GetAvailableDriveSpace = "Error"
    Err = 0
Else
    If Val(sSpace) < 1000000 Then
        GetAvailableDriveSpace = Format$(sSpace, "###,###,###,###,##0 KB")
    ElseIf Val(sSpace) > 999999 And Val(sSpace) < 1000000 Then
        GetAvailableDriveSpace = Format$(sSpace, "###,###,###,###,##0 MB")
    ElseIf Val(sSpace) > 9999999 Then
        GetAvailableDriveSpace = Format$(sSpace, "###,###,###,###,##0 GB")
    End If
End If

End Function
Function OpenTextFile(cControl As TextBox, sFileName As String) As Boolean

'This function opens a specified text file into the calling TextBox...

On Local Error GoTo OpenTextFileError

Dim sNam As String
Dim FileFree As Long

'Read in the file...
FileFree = FreeFile
Open sFileName For Input As #FileFree
    Do
        Line Input #FileFree, sNam
        cControl = cControl & sNam
    Loop Until EOF(FileFree)
Close #FileFree

Exit Function



OpenTextFileError:
    Close 'Close ALL open files...
    Call WriteToErrorLog("Global", "OpenTextFileError", Err.Description, Err.Number, True)
    Exit Function

End Function
Function OpenDB(DB As Database, sPassWord As String) As Boolean

'This routine waits until the db is free and then opens it for the user.
'It times out when QuickRef.DBTimeOut reaches it's limit and returns an open error...

On Local Error Resume Next

Dim iOpenCount As Long

Do
    Err.Number = 0
    Set DB = OpenDatabase(QuickRef.DBFileName, True, False, ";pwd=" & sPassWord)
    iOpenCount = iOpenCount + 1
Loop Until Err.Number = 0 Or iOpenCount >= QuickRef.DBTimeOut

'Set function to true or false...
If iOpenCount >= QuickRef.DBTimeOut Then
    OpenDB = False
Else
    OpenDB = True
End If

End Function
Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Function StartExternalProgram(sProgramPathAndName As String) As Boolean

'This function will attempt to launch an external program, such as MS-Word for example...

On Local Error GoTo StartWordError

'Shell the application...
Call Shell("Start " & sProgramPathAndName, vbMaximizedFocus)

StartExternalProgram = True
Exit Function



StartWordError:
    Call WriteToErrorLog("GlobalFunctions", "StartExternalProgramError", Err.Description, Err.Number, True)
    Exit Function

End Function
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean

On Local Error Resume Next

Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)

WriteINI = (Err.Number = 0)

End Function
Function ReadINI(sSection As String, sKeyName As String, sINIFileName As String) As String

On Local Error Resume Next

Dim sRet As String

sRet = String(255, Chr(0))

'Note: INI Filename can point to a local ini file or a remote ini file...
ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))

End Function
Sub WriteToErrorLog(sFormName As String, sRoutineName As String, sError As String, iErrorNumber As Long, iDisplayMsgBox As Boolean, Optional sMsg As String, Optional iMsgIcon As Integer)

'This sub routine writes errors that occur throughout the entire program to a text file so that we, (the programmers),
'can know what errors are occuring in the system...

'Temporarily set to true so I can see all errors that are occuring in the program...jeffdeaton
iDisplayMsgBox = True

On Local Error Resume Next

Dim FileFree As Integer

'Append this error to the ErrorLog.Txt file...
FileFree = FreeFile
Open App.Path & "\ErrorLog.Txt" For Append As #FileFree
    Print #FileFree, sFormName, sRoutineName, sError, iErrorNumber
Close #FileFree

'if iDisplayMsgBox = True then display the error that just happened to the user...
If iDisplayMsgBox = True Then
    'Default the MsgBox icon to an Information icon...
    If iMsgIcon = 0 Then
        iMsgIcon = vbCritical
    End If
    'No custom text msg was brought in so use the default message...
    If Trim$(sMsg) = "" Then
        MsgBox "The following error has occured in your program: " & vbCrLf & vbCrLf & sError & vbCrLf & vbCrLf & "Error Number: " & iErrorNumber, iMsgIcon, "Error..."
    'Use a custom message...
    Else
        MsgBox sMsg, iMsgIcon, "Error..."
    End If
End If

End Sub
