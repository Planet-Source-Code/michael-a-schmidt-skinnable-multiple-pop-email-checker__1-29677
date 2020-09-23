VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Tag             =   "SKIPSKIN"
   Begin VB.PictureBox picAuthor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   30
      Picture         =   "frmAbout.frx":014A
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   0
      Width           =   1530
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "~ Thanks to Jeff Deaton for the original Skin Engine ~"
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   1890
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ":: Open Source ::"
      Height          =   225
      Left            =   1260
      TabIndex        =   8
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ":: Freeware ::"
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   1650
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ":: December 5th, 2001 ::"
      Height          =   225
      Left            =   2520
      TabIndex        =   6
      Top             =   1650
      Width           =   2025
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " - Seperate Mail Instancing"
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   3165
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " - Encrypted Passwords"
      Height          =   225
      Left            =   1800
      TabIndex        =   4
      Top             =   1170
      Width           =   3165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - Multiple Account Support"
      Height          =   225
      Left            =   1800
      TabIndex        =   3
      Top             =   750
      Width           =   3165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "©2001 Mike Schmidt"
      Height          =   195
      Left            =   1260
      TabIndex        =   1
      Top             =   330
      Width           =   3165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """Neko Email"""
      Height          =   165
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
