VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Login to Contract Control"
   ClientHeight    =   2325
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   3345
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0442
   ScaleHeight     =   1373.688
   ScaleMode       =   0  'User
   ScaleWidth      =   3140.774
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUserName 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Enter User name or select one from list"
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter Password"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Control Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   233
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   742
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   1102
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPassRS As ADODB.Recordset
Attribute adoPassRS.VB_VarHelpID = -1


Private Sub cboUserName_LostFocus()
UserName = cboUserName.Text
End Sub

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim FindUser As String

FindUser = cboUserName.Text
Set adoPassRS = New ADODB.Recordset
adoPassRS.Open "Select * From UserList where user = '" & FindUser & "'", DB, adOpenStatic, adLockOptimistic
If Not adoPassRS.BOF Then
   adoPassRS.MoveFirst
End If
If adoPassRS.RecordCount <> 0 Then
If adoPassRS("Password") = txtPassword.Text Then
 LoginSucceeded = True
 Me.Hide
ElseIf adoPassRS("password") <> txtPassword.Text Then
 MsgBox "Incorrect Password"
 txtPassword.Text = ""
 txtPassword.SetFocus
End If
Else
 MsgBox "User Does Not Exits"
 cboUserName.Text = ""
 cboUserName.SetFocus
End If
End Sub
Private Sub Form_Load()
SetRegion
 Set adoPassRS = New ADODB.Recordset
 adoPassRS.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
    If adoPassRS.RecordCount = 0 Then
        LoginSucceeded = True
    Else
        adoPassRS.MoveFirst
        Do While Not adoPassRS.EOF
        cboUserName.AddItem (adoPassRS("User"))
            '.AddItem IIf(IsNull(adoPrimaryRSPass("User_Name")), "", adoPrimaryRSPass("User_Name"))
        adoPassRS.MoveNext
        Loop
    End If
End Sub
Private Sub SetRegion()
    On Error Resume Next
    If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(Me.Picture, RGB(255, 0, 255))
    SetWindowRgn Me.hwnd, hRgn, True
End Sub
