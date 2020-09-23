VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddSubCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontractor Details"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmAddSubCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtSubConNum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ListView lstSubCon 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   6376
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name :"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Name :"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Number :"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddSubCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SubConRs As ADODB.Recordset
Dim LSubRs As ADODB.Recordset
Dim MaxNum As Integer

Private Sub cmdAddNew_Click()
Dim NumRs As ADODB.Recordset
Dim sqlStatement As String


If cmdAddNew.Caption = "Add New" Then
 If Not DEnv.rsSubConMaxNum.State = adStateClosed Then DEnv.rsSubConMaxNum.Close
 sqlStatement = "SELECT max(subConNum) as MaxNum FROM SubConDetails"
 DEnv.rsSubConMaxNum.Open sqlStatement, DB, adOpenStatic, adLockOptimistic
 DEnv.rsSubConMaxNum.Requery
 If DEnv.rsSubConMaxNum!MaxNum <> 0 Then
  MaxNum = DEnv.rsSubConMaxNum!MaxNum + 1
 Else
  MaxNum = 1
 End If
   Set SubConRs = New ADODB.Recordset
   SubConRs.Open "Select * FROM SubConDetails", DB, adOpenStatic, adLockOptimistic
   Set txtFirstName.DataSource = SubConRs
    txtFirstName.DataField = "SubConFname"
   Set txtLastName.DataSource = SubConRs
    txtLastName.DataField = "SubConLname"
    SubConRs.AddNew
     SubConRs!SubConNum = MaxNum
     txtSubConNum.Text = SubConRs!SubConNum
     txtFirstName.SetFocus
     cmdAddNew.Caption = "Save New"
ElseIf cmdAddNew.Caption = "Save New" Then
 cmdAddNew.Caption = "Add New"

    SubConRs.Update
    SubConRs.Close
    OpenSubConRs
    DoList
End If
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me

End Sub

Private Sub cmdDelete_Click()
Dim Response As Integer
    If Not SubConRs.EOF Or SubConRs.BOF Then
          Response = MsgBox("Sure?", vbQuestion + vbYesNo)
        If Response = vbYes Then
        
         SubConRs.Delete
        ElseIf Response = vbNo Then
        End If
    End If
     OpenSubConRs
     DoList
End Sub

Private Sub cmdPrint_Click()
rptAllSubCon.Show
End Sub

Private Sub cmdUpdate_Click()
SubConRs.Update
SubConRs.Close
OpenSubConRs
DoList
End Sub

Private Sub Form_Load()
DoList
OpenSubConRs

End Sub

Private Sub OpenSubConRs()
Set SubConRs = New ADODB.Recordset
SubConRs.Open "Select * FROM SubConDetails", DB, adOpenStatic, adLockOptimistic
Set txtSubConNum.DataSource = SubConRs
txtSubConNum.DataField = "SubConNum"
Set txtFirstName.DataSource = SubConRs
txtFirstName.DataField = "SubConFname"
Set txtLastName.DataSource = SubConRs
txtLastName.DataField = "SubConLname"
End Sub
Private Sub DoList()
Dim itmx As ListItem
lstSubCon.ColumnHeaders.Add , , "First Name, Last Name", lstSubCon.Width - 100
lstSubCon.ColumnHeaders.Add , , "SubConNum", 1

Set LSubRs = New ADODB.Recordset
    LSubRs.Open "SELECT * FROM SubconDetails", DB, adOpenStatic, adLockOptimistic
    
    If Not LSubRs.BOF Then LSubRs.MoveFirst
    lstSubCon.ListItems.Clear
    Do While Not LSubRs.EOF
        Set itmx = lstSubCon.ListItems.Add(, , LSubRs!subconfname & ", " & LSubRs!subconLname)
        itmx.SubItems(1) = LSubRs!SubConNum
        LSubRs.MoveNext
    Loop
    If Not LSubRs.EOF Then LSubRs.MoveFirst
    lstSubCon.Refresh
 If LSubRs.RecordCount = 0 Then
  cmdUpdate.Enabled = False
  cmdDelete.Enabled = False
 ElseIf LSubRs.RecordCount <> 0 Then
  cmdUpdate.Enabled = True
  cmdDelete.Enabled = True
 End If
End Sub

Private Sub FromListUpdateRecord()
    On Error GoTo ExiThis
    If Not SubConRs.BOF Then SubConRs.MoveFirst
    If Not lstSubCon.SelectedItem.Text = Empty Then
        SubConRs.Find "SubConNum='" & Trim(lstSubCon.SelectedItem.SubItems(1)) & "'"
    End If
ExiThis:
End Sub

Private Sub lstSubCon_Click()
FromListUpdateRecord
End Sub
Private Sub lstSubCon_KeyUp(KeyCode As Integer, Shift As Integer)
FromListUpdateRecord
End Sub

Private Sub txtFirstName_LostFocus()
If txtFirstName = "" Then
 MsgBox "Must supply a First Name"
 txtFirstName.SetFocus
End If
End Sub
