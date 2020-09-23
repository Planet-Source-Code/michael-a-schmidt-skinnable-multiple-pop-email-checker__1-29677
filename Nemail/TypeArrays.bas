Attribute VB_Name = "TypeArrays"
Option Explicit

'Global Type Arrays...
Global Help As tHelp
Global Login As tLogin
Global QuickRef As tQuickRef
Global FS As New FileSystemObject

'Quick Reference...
Type tQuickRef
    CancelOperation As Boolean
    DBPassWord As String
    DBFileName As String
    DBTimeOut As Long
    UserINIFileName As String
    GlobalINIFileName As String
End Type

'Login...
Type tLogin
    UserID As String
    LoginID As Long
    UserFullName As String
    Administrator As Boolean
    LoginDateAndTime As Date
    LogoutDateAndTime As Date
    ReLoggingIn As Boolean
End Type

'Tech Support...
Type tHelp
    TechSupportCompany As String
    TechSupportPhone1 As String
    TechSupportPhone2 As String
    TechSupportPhone3 As String
    TechSupportFax As String
End Type
