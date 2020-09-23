VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubConReportSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Select"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmSubConReportSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmrptSelection 
      Caption         =   "Type of Reports"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optPaymentrpt 
         Caption         =   "Payments Report"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton optWorkDetailsrpt 
         Caption         =   "Work Details Report"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.ComboBox cboContractNo 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboSubCon 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   24444929
      CurrentDate     =   37150
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Date :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Contract Numbers :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Subcontractors :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmSubConReportSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubConRs As ADODB.Recordset
Dim ConNoRs As ADODB.Recordset
Dim SubConNum As Integer
Dim ContractNo As String
Dim SDate As String
Dim sqlString As String

Private Sub cboSubCon_Click()
'SubConNum = 0
'SubConNum = cboSubCon.ItemData(cboSubCon.ListIndex)
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdPrint_Click()
SubConNum = 0
If cboSubCon.Text <> "" Then
SubConNum = cboSubCon.ItemData(cboSubCon.ListIndex)
End If '.text <> ""
ContractNo = cboContractNo.Text
If Not IsDate(dtpDate.Value) Then
SDate = Empty
Else
SDate = Format(dtpDate.Value, "dd/mm/yyyy")
End If 'Isdate
If SubConNum = 0 Then
 If optWorkDetailsrpt.Value = True Then
  If Not DEnv.rscmdSubCon.State = adStateClosed Then DEnv.rscmdSubCon.Close
   sqlString = "SELECT DISTINCT SubConDetails.* FROM SubConDetails"
   DEnv.rscmdSubCon.Open sqlString, DB, adOpenStatic, adLockOptimistic
   DEnv.rscmdSubCon.Filter = ""
   DEnv.rscmdSubCon.Requery
   rptWorkDetails.Show
   DEnv.rscmdSubCon.Close
 ElseIf optPaymentrpt.Value = True Then
  If Not DEnv.rscmdPayments.State = adStateClosed Then DEnv.rscmdPayments.Close
   sqlString = "SELECT DISTINCT SubConDetails.* FROM SubConDetails, SubConPayments WHERE SubConDetails.SubConNum = SubConPayments.SubConNum"
   DEnv.rscmdPayments.Open sqlString, DB, adOpenStatic, adLockOptimistic
   DEnv.rscmdPayments.Filter = ""
   DEnv.rscmdPayments.Requery
   rptSubPayments.Show
   DEnv.rscmdPayments.Close
 End If '.Value = True

ElseIf SubConNum <> 0 Then
 If optWorkDetailsrpt.Value = True Then
  If Not DEnv.rscmdSubCon.State = adStateClosed Then DEnv.rscmdSubCon.Close
  sqlString = "SELECT DISTINCT SubConDetails.* FROM SubConDetails, SubConWorkDetails WHERE " & _
  "SubConDetails.SubConNum = SubConWorkDetails.SubConNum " & _
  "AND (SubConWorkDetails.ContractNo <> '')"
  DEnv.rscmdSubCon.Open sqlString, DB, adOpenStatic, adLockOptimistic
  DEnv.rscmdSubCon.Filter = "SubConNum=" & SubConNum
  DEnv.rscmdSubCon.Requery
   If DEnv.rscmdSubCon.RecordCount <> 0 Then
    rptWorkDetails.Show
   End If 'Recordcount
  DEnv.rscmdSubCon.Close
 ElseIf optPaymentrpt.Value = True Then
  If Not DEnv.rscmdPayments.State = adStateClosed Then DEnv.rscmdPayments.Close
  sqlString = "SELECT DISTINCT SubConDetails.* FROM SubConDetails, SubConPayments WHERE SubConDetails.SubConNum = SubConPayments.SubConNum"
  DEnv.rscmdPayments.Open sqlString, DB, adOpenStatic, adLockOptimistic
  DEnv.rscmdPayments.Filter = "SubConNum=" & SubConNum
  DEnv.rscmdPayments.Requery
   If DEnv.rscmdPayments.RecordCount <> 0 Then
    rptSubPayments.Show
   End If 'Recordcount
  DEnv.rscmdPayments.Close
 End If '.value = true
End If 'SubconNum
End Sub


Private Sub Form_Load()
Dim r As Integer
Dim ItmX As ListItem
cboSubCon.Text = ""
cboContractNo.Text = ""
dtpDate.Value = Now
dtpDate.Value = ""
Set SubConRs = New ADODB.Recordset
Set ConNoRs = New ADODB.Recordset
sqlString = "Select * from SubconDetails"
SubConRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
If SubConRs.RecordCount <> 0 Then
SubConRs.MoveFirst
 With SubConRs
   For r = 1 To SubConRs.RecordCount
   cboSubCon.AddItem (!subconfname & "," & !subconLname)
   cboSubCon.ItemData(cboSubCon.NewIndex) = !SubConNum
   .MoveNext
   Next ' 1 to subconrs.recordcount
 End With 'SubconRs
End If 'SubCOnrs.RecordCount
sqlString = "Select ContractNo,ClientName from contractdetails"
ConNoRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
If ConNoRs.RecordCount <> 0 Then
ConNoRs.MoveFirst
 With ConNoRs
  For r = 1 To ConNoRs.RecordCount
  cboContractNo.AddItem (!ContractNo)
  .MoveNext
  Next ' 1 to connors.recordcount
 End With 'connors
End If 'Connors.recordcount

End Sub

