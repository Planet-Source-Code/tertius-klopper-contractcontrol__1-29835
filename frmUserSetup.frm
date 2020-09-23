VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Setup"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmUserSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtReenter 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "Delete User"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "New User"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "Enter User Name"
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3635
      EndProperty
   End
   Begin VB.Label lblNewPassword 
      Caption         =   "New Password :"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblOldPassword 
      Caption         =   "Old Password :"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblReenter 
      Caption         =   "Reenter Password :"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password :"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Name :"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmUserSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LUsersRs As ADODB.Recordset
Dim UserRs As ADODB.Recordset
Dim OldPassword As String
Private Sub cmdAddUser_Click()
Dim UserXRs As ADODB.Recordset
    'If Not CusRS.State = adStateClosed Then CusRS.Close
If cmdAddUser.Caption = "New User" Then
    lblReenter.Visible = True
    txtReenter.Visible = True
    Set UserRs = New ADODB.Recordset
    UserRs.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
    Set txtUserName.DataSource = UserRs
    txtUserName.DataField = "User"
    Set txtPassword.DataSource = UserRs
    txtPassword.DataField = "Password"
    UserRs.AddNew
    txtUserName.SetFocus
    cmdAddUser.Caption = "Save User"
ElseIf cmdAddUser.Caption = "Save User" Then
'      On Error GoTo errFucks
    If txtUserName.Text = Empty Or txtPassword.Text = Empty Then
        MsgBox "Must supply a User Name and/or Password", vbCritical
        If txtUserName.Text = Empty Then txtUserName.SetFocus
        If txtUserName.Text <> Empty And txtPassword.Text = Empty Then txtPassword.SetFocus
        Exit Sub
    End If
    Set UserXRs = New ADODB.Recordset
    UserXRs.Open "SELECT * FROM userlist where user = '" & txtUserName.Text & "'", DB, adOpenStatic, adLockOptimistic
   If UserXRs.RecordCount = 0 Then
    If txtPassword.Text = txtReenter.Text Then
      UserRs.Update
      UserRs.Close
      cmdAddUser.Caption = "New User"
      lblReenter.Visible = False
      txtReenter.Visible = False
      txtReenter.Text = ""
      Set UserRs = New ADODB.Recordset
      UserRs.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
      Set txtUserName.DataSource = UserRs
      txtUserName.DataField = "User"
      Set txtPassword.DataSource = UserRs
      txtPassword.DataField = "Password"
      DoList
      Exit Sub
    ElseIf txtPassword.Text <> txtReenter.Text Then
      MsgBox "Passwords does not match, Try Again", vbCritical
      txtPassword.Text = ""
      txtReenter.Text = ""
      txtPassword.SetFocus
    End If
   ElseIf UserXRs.RecordCount <> 0 Then
     MsgBox "User Name already Exits, Try another User Name", vbCritical
     txtUserName.Text = ""
     txtPassword.Text = ""
     txtReenter.Text = ""
     txtUserName.SetFocus
   End If 'userxrs.recordcount
'errFucks:
'    MsgBox "oops! Unexpacted Error, contact vendor."
End If
End Sub

Private Sub cmdChangePassword_Click()
Dim UserCRS As ADODB.Recordset


If cmdChangePassword.Caption = "Change Password" Then
    cmdChangePassword.Caption = "Save New Password"
    ChangePos
    OldPassword = txtPassword.Text
    txtOldPassword.SetFocus
    txtPassword.Text = ""
ElseIf cmdChangePassword.Caption = "Save New Password" Then
 If txtPassword.Text = txtReenter.Text And txtPassword <> "" Then
  cmdChangePassword.Caption = "Change Password"
  
  
  UserRs.Update
  UserRs.Close
  StartPos
    txtOldPassword.Text = ""
    txtReenter.Text = ""
    Set UserRs = New ADODB.Recordset
    UserRs.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
    Set txtUserName.DataSource = UserRs
    txtUserName.DataField = "User"
    Set txtPassword.DataSource = UserRs
    txtPassword.DataField = "Password"
    
    DoList
 ElseIf txtPassword.Text <> txtReenter.Text Or txtPassword = "" Then
    MsgBox "Password mismatch or Password empty", , "Password Error"
    txtPassword.SetFocus
 End If
End If

End Sub

Private Sub cmdDeleteUser_Click()
    Dim Response As Integer
    If Not UserRs.EOF Or UserRs.BOF Then
        Response = MsgBox("Sure?", vbQuestion + vbYesNo)
        If Response = vbYes Then
         UserRs.Delete
         txtUserName.Text = ""
         txtPassword.Text = ""
         txtReenter.Text = ""
        ElseIf Response = vbNo Then
        
        End If
    End If
      Set UserRs = New ADODB.Recordset
      UserRs.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
      Set txtUserName.DataSource = UserRs
      txtUserName.DataField = "User"
      Set txtPassword.DataSource = UserRs
      txtPassword.DataField = "Password"
     DoList
End Sub

Private Sub Form_Load()
Set UserRs = New ADODB.Recordset
UserRs.Open "Select * from UserList", DB, adOpenStatic, adLockOptimistic
Set txtUserName.DataSource = UserRs
txtUserName.DataField = "User"
Set txtPassword.DataSource = UserRs
txtPassword.DataField = "Password"
DoList
End Sub
Private Sub DoList()

Set LUsersRs = New ADODB.Recordset
LUsersRs.Open "Select user from UserList order by User", DB, adOpenStatic, adLockOptimistic
lstUsers.ListItems.Clear
If Not LUsersRs.BOF Then LUsersRs.MoveFirst
Do While Not LUsersRs.EOF
    lstUsers.ListItems.Add , , LUsersRs("User")
    LUsersRs.MoveNext
Loop
If LUsersRs.RecordCount = 0 Then
   cmdDeleteUser.Enabled = False
   cmdChangePassword.Enabled = False
ElseIf LUsersRs.RecordCount <> 0 Then
   cmdDeleteUser.Enabled = True
   cmdChangePassword.Enabled = True
End If
lstUsers.Refresh
End Sub
Private Sub lstUsers_Click()
  FromListUpdate
End Sub

Private Sub FromListUpdate()
    On Error GoTo ExiThis
    If Not UserRs.BOF Then UserRs.MoveFirst
    If Not lstUsers.SelectedItem.Text = Empty Then
        UserRs.Find "User='" & Trim(lstUsers.SelectedItem.Text) & "'"
    End If
ExiThis:
End Sub

Private Sub txtOldPassword_LostFocus()
If OldPassword <> txtOldPassword.Text Then
  MsgBox "Incorrect Password", , "Password Error"
  txtOldPassword.SetFocus
End If
End Sub

Private Sub StartPos()
    cmdAddUser.Enabled = True
    cmdDeleteUser.Enabled = True
    lstUsers.TabIndex = 0
    txtUserName.TabIndex = 1
    txtPassword.TabIndex = 2
    txtReenter.TabIndex = 3
    txtOldPassword.TabIndex = 4
    cmdAddUser.TabIndex = 6
    cmdDeleteUser.TabIndex = 7
    cmdChangePassword.TabIndex = 8
    lblPassword.Top = 600
    txtPassword.Top = 600
    lblReenter.Top = 960
    txtReenter.Top = 960
    lblOldPassword.Top = 1320
    txtOldPassword = 1320
    lblNewPassword.Top = 1680
    lstUsers.Enabled = True
    txtUserName.Enabled = True
    lblPassword.Visible = True
    txtPassword.Visible = True
    lblNewPassword.Visible = False
    lblReenter.Visible = False
    txtReenter.Visible = False
    lblOldPassword.Visible = False
    txtOldPassword.Visible = False
End Sub

Private Sub ChangePos()
    cmdAddUser.Enabled = False
    cmdDeleteUser.Enabled = False
    lblPassword.Visible = False
    txtPassword.Visible = True
    txtPassword.Top = 960
    lblNewPassword.Top = 960
    lblReenter.Top = 1320
    txtReenter.Top = 1320
    lblOldPassword.Top = 600
    txtOldPassword.Top = 600
    txtOldPassword.TabIndex = 1
    txtPassword.TabIndex = 2
    txtReenter.TabIndex = 3
    lstUsers.Enabled = False
    txtUserName.Enabled = False
    lblNewPassword.Visible = True
    lblReenter.Visible = True
    txtReenter.Visible = True
    lblOldPassword.Visible = True
    txtOldPassword.Visible = True

End Sub
