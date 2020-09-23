VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   4515
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   7515
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3116.333
   ScaleMode       =   0  'User
   ScaleWidth      =   7056.974
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      DrawMode        =   15  'Merge Pen Not
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H8000000A&
      Height          =   540
      Left            =   1170
      Picture         =   "frmAbout.frx":6DCEE
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5175
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label lblOsVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblOsVersion"
      Height          =   255
      Left            =   1980
      TabIndex        =   8
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "lblUserName"
      Height          =   255
      Left            =   1980
      TabIndex        =   7
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   986.004
      X2              =   6056.884
      Y1              =   2153.479
      Y2              =   2153.479
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 Tertius Klopper"
      Height          =   255
      Left            =   1980
      TabIndex        =   6
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6E5B8
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   1980
      TabIndex        =   2
      Top             =   2160
      Width           =   4965
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1980
      TabIndex        =   4
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblProgramVersion"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1980
      TabIndex        =   5
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6E693
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   1185
      TabIndex        =   3
      Top             =   3240
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SetRegion()
    On Error Resume Next
    If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(Me.Picture, RGB(255, 0, 255))
    SetWindowRgn Me.hwnd, hRgn, True
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
     SetRegion
     On Error Resume Next
     Me.Caption = "About " & App.Title
     lblVersion.Caption = "DEMO Version " & App.Major & "." & App.Minor & "." & App.Revision
     lblTitle.Caption = App.Title
     lblUserName.Caption = "Windows Register To: " & _
     GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "RegisteredOwner")

     lblOsVersion.Caption = "Version: " & _
     GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "ProductName") & " - " & _
     GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "VersionNumber")
End Sub

