VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddContract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract Details"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "frmAddContract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddContract 
      Caption         =   "New Contract"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chkComplete 
      Caption         =   "Contract Completed"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      Height          =   975
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtWorkDesc 
      Height          =   765
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txtClientName 
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtContractNo 
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtEstimator 
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtAmount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   7177
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtQuoteNo 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ListView lstContracts 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label8 
      Caption         =   "Client Address :"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Work Description :"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Client Name :"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Contract No :"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Estimator :"
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Amount :"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Quote No :"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date - DD/MM/YYYY"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConRs As ADODB.Recordset
Dim LConRs As ADODB.Recordset
Dim sqlString As String
Private Sub cmdAddContract_Click()
 If cmdAddContract.Caption = "New Contract" Then
   OpenConRs
    ConRs.AddNew
     txtDate.SetFocus
     cmdAddContract.Caption = "Save Contract"
     cmdDelete.Enabled = False
     cmdUpdate.Enabled = False
ElseIf cmdAddContract.Caption = "Save Contract" Then
 cmdAddContract.Caption = "New Contract"
     cmdDelete.Enabled = True
     cmdUpdate.Enabled = True
'     txtDate = mskDate.Text
     ConRs.Update
     ConRs.Close
     OpenConRs
     DoList
End If
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Response As Integer
    If Not ConRs.EOF Or ConRs.BOF Then
          Response = MsgBox("Sure?", vbQuestion + vbYesNo)
        If Response = vbYes Then
        
         ConRs.Delete
        ElseIf Response = vbNo Then
        End If
    End If
     OpenConRs
     DoList
End Sub

Private Sub cmdPrint_Click()
frmConReportSelect.OptConList.Value = True
frmConReportSelect.Show
End Sub

Private Sub cmdUpdate_Click()
ConRs.Update
ConRs.Close
OpenConRs
DoList
End Sub

Private Sub Form_Load()
OpenConRs
DoList
End Sub
Private Sub OpenConRs()
Set ConRs = New ADODB.Recordset
ConRs.Open "Select * FROM ContractDetails", DB, adOpenStatic, adLockOptimistic
Set txtDate.DataSource = ConRs
txtDate.DataField = "Date"
Set txtQuoteNo.DataSource = ConRs
txtQuoteNo.DataField = "QuoteNo"
Set txtAmount.DataSource = ConRs
txtAmount.DataField = "Amount"
Set txtEstimator.DataSource = ConRs
txtEstimator.DataField = "Estimator"
Set txtContractNo.DataSource = ConRs
txtContractNo.DataField = "ContractNo"
Set txtClientName.DataSource = ConRs
txtClientName.DataField = "ClientName"
Set txtWorkDesc.DataSource = ConRs
txtWorkDesc.DataField = "WorkDescription"
Set txtAddress.DataSource = ConRs
txtAddress.DataField = "ClientAddress"
Set chkComplete.DataSource = ConRs
chkComplete.DataField = "Completed"
End Sub

Private Sub DoList()
Dim ItmX As ListItem
lstContracts.ColumnHeaders.Add , , "Contract Num", lstContracts.Width / 2.95
lstContracts.ColumnHeaders.Add , , "Client Name", lstContracts.Width
    Set LConRs = New ADODB.Recordset
    LConRs.Open "SELECT * FROM ContractDetails", DB, adOpenStatic, adLockOptimistic
    If Not LConRs.BOF Then LConRs.MoveFirst
    lstContracts.ListItems.Clear
    Do While Not LConRs.EOF
        Set ItmX = lstContracts.ListItems.Add(, , LConRs!ContractNo)
         ItmX.SubItems(1) = LConRs!Clientname
        LConRs.MoveNext
    Loop
    If Not LConRs.EOF Then LConRs.MoveFirst
    lstContracts.Refresh
 If LConRs.RecordCount = 0 Then
  cmdUpdate.Enabled = False
  cmdDelete.Enabled = False
 ElseIf LConRs.RecordCount <> 0 Then
  cmdUpdate.Enabled = True
  cmdDelete.Enabled = True
 End If
End Sub
Private Sub FromListUpdateRecord()
On Error GoTo ExiThis
    If Not ConRs.BOF Then ConRs.MoveFirst
    If Not lstContracts.SelectedItem.Text = Empty Then
        ConRs.Find "ContractNo='" & Trim(lstContracts.SelectedItem.Text) & "'"
    End If
ExiThis:
End Sub
Private Sub lstContracts_Click()
 FromListUpdateRecord
End Sub

Private Sub txtAmount_GotFocus()
txtAmount.Text = ""
End Sub

Private Sub txtAmount_LostFocus()
If txtAmount.Text = "" Then
  MsgBox "Must Supply a Amount may not be 0"
  txtAmount.SetFocus
End If
End Sub

Private Sub txtClientName_LostFocus()
If txtClientName.Text = "" And txtContractNo <> "" Then
   MsgBox "Must Supply a Client Name"
   txtClientName.SetFocus
End If
End Sub

Private Sub txtContractNo_LostFocus()
Dim ConNum As ADODB.Recordset
Dim sqlString As String

If txtContractNo.Text = "" Then
   MsgBox "Must Supply a Contract No"
   txtContractNo.SetFocus
End If
If txtContractNo.Text <> "" Then
 Set ConNum = New ADODB.Recordset
 sqlString = "Select contractno from contractdetails  where contractno ='" & Trim(txtContractNo.Text) & "'"
 ConNum.Open sqlString, DB, adOpenStatic, adLockOptimistic
   If ConNum.RecordCount <> 0 Then
      MsgBox "The Contract number " & txtContractNo.Text & " already exist"
      txtContractNo.Text = ""
      txtContractNo.SetFocus
   End If
End If
End Sub


Private Sub txtDate_LostFocus()
If txtDate.Text = "" Then
 MsgBox "Must Supply a Date"
 txtDate.SetFocus
ElseIf Not IsDate(txtDate.Text) Then
 MsgBox "Not a valid Date"
 txtDate.SetFocus
End If
End Sub
