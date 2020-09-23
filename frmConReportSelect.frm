VERSION 5.00
Begin VB.Form frmConReportSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Select"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmConReportSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboContractNo 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   1410
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2265
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   585
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame fraTypeofRpt 
      Caption         =   "Type of Reports"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton optConDetails 
         Caption         =   "Contract Details"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton OptConList 
         Caption         =   "Contract List"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Contract Number :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmConReportSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConNoRs As ADODB.Recordset
Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdPrint_Click()
If OptConList.Value = True Then
 rptConList.Show
ElseIf optConDetails.Value = True Then
If cboContractNo.Text <> "" Then
  If Not DEnv.rscmdConDetails.State = adStateClosed Then DEnv.rscmdConDetails.Close
  sqlString = "SELECT ContractDetails.* FROM ContractDetails"
  DEnv.rscmdConDetails.Open sqlString, DB, adOpenStatic, adLockOptimistic
  DEnv.rscmdConDetails.Filter = "ContractNo='" & cboContractNo.Text & "'"
  DEnv.rscmdConDetails.Requery
  rptConDetails.Show
  DEnv.rscmdConDetails.Close
ElseIf cboContractNo.Text = "" Then
 MsgBox "Must Supply a Contract Number", vbInformation
End If 'cboContractNo.text
End If '.value = true
End Sub

Private Sub Form_Load()
Set ConNoRs = New ADODB.Recordset
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
