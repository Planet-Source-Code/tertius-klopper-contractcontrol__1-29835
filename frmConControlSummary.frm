VERSION 5.00
Begin VB.Form frmConControlSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract Control Summary"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmConControlSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   2078
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   518
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblProfitPer 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Profit Margin :"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1800
      X2              =   3840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblOutStanding 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Amount Outstanding :"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblProfit 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblTotalExp 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblTotalPay 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Profit :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Total Expences :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Total Payments :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblContractAmount 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Contract Amount :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblClientName 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblContractNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Client Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Contract Number :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmConControlSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim ConNum As String
If Not DEnv.rscmdConConSum.State = adStateClosed Then DEnv.rscmdConConSum.Close
   sqlString = "SELECT DISTINCT ContractNo, ClientName, Amount FROM ContractDetails"
   DEnv.rscmdConConSum.Open sqlString, DB, adOpenStatic, adLockOptimistic
   DEnv.rscmdConConSum.Filter = "ContractNo='" & frmConControlSummary.lblContractNo.Caption & "'"
   DEnv.rscmdConConSum.Requery
   rptConConSummary.Show
   DEnv.rscmdConConSum.Close
End Sub

Private Sub Form_Activate()
DoCalc
End Sub

Private Sub DoCalc()
Dim ConAmount As Currency
Dim OutStand As Currency
Dim Profit As Currency
Dim ProfitPer As Single
ConAmount = frmContractControl.lstContracts.SelectedItem.SubItems(2)

OutStand = ConAmount - TotalPay
Profit = TotalPay - TotalExp
If TotalPay <> 0 Then
ProfitPer = Profit / TotalPay
Else
ProfitPer = 0
End If
lblContractNo.Caption = frmContractControl.lstContracts.SelectedItem
lblClientName.Caption = frmContractControl.lstContracts.SelectedItem.SubItems(1)
lblContractAmount.Caption = FormatCurrency(ConAmount, 2)
lblOutStanding.Caption = FormatCurrency(OutStand, 2)
lblTotalPay.Caption = FormatCurrency(TotalPay, 2)
lblTotalExp.Caption = FormatCurrency(TotalExp, 2)
If Profit >= 0 Then
lblProfit.ForeColor = &HFF0000
Else
lblProfit.ForeColor = &HFF&
End If
lblProfit.Caption = FormatCurrency(Profit, 2)
lblProfitPer.Caption = FormatPercent(ProfitPer, 2)
End Sub

