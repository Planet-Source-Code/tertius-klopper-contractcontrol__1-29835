VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Work Details"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgExport 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   3113
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   1553
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Press Start to Begin Export Process"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubConNum As Integer
Dim SubConRs As ADODB.Recordset
Dim SubConDueRs As ADODB.Recordset
Dim ConConRS As ADODB.Recordset
Dim ConExpRs As ADODB.Recordset
Dim SubWorkRs As ADODB.Recordset
Dim sqlString As String
Private Sub SetRegion()
    On Error Resume Next
    Me.Picture = frmImages.Pic2.Picture
    If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(Me.Picture, RGB(255, 0, 255))
    SetWindowRgn Me.hwnd, hRgn, True
End Sub

Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdStart_Click()
prgExport.Value = 0
 
 If MsgBox("Process cannot be stoped ones it has started!" & Chr(13) & "Contunue?", vbInformation + vbYesNo, "Export New Records") = vbYes Then
  cmdCancel.Enabled = False
   lblStatus.Caption = "Adding New Payments Records"
   AddNewSubCon
   lblStatus.Caption = "Adding New Contract Expences"
   AddNewConCon
   lblStatus.Caption = "Adding Export Time Stamp"
   TimeStamp
   lblStatus.Caption = "Export Complete"
 End If
cmdCancel.Enabled = True
End Sub

Private Sub AddNewSubCon()
Dim r As Integer
Dim TodayD As String
Dim SubConGetRs As ADODB.Recordset
Dim SubConPayRs As ADODB.Recordset
Dim CFA As Currency
prgExport.Value = 0
TodayD = Date
  Set SubConRs = New ADODB.Recordset
  sqlString = "Select subconnum from subcondetails"
  SubConRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
  If SubConRs.RecordCount <> 0 Then
   SubConRs.MoveFirst
   With SubConRs
     prgExport.Max = SubConRs.RecordCount
     For r = 1 To SubConRs.RecordCount
       Set SubConDueRs = New ADODB.Recordset
       sqlString = "SELECT sum(Amount) AS TotalDue From SubConWorkDetails " & _
       "where SubConNum=" & !SubConNum & _
       " and TransferDate is null"
       SubConDueRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
       If SubConDueRs.RecordCount <> 0 And SubConDueRs!TotalDue <> 0 Then
        With SubConDueRs
         Set SubConGetRs = New ADODB.Recordset
         sqlString = "Select * from SubConPayments " & _
         "where SubconNum=" & SubConRs!SubConNum & _
         " and Date=#" & Format(TodayD, "m/d/yyyy") & "#"
         SubConGetRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
         If SubConGetRs.RecordCount <> 0 Then
           With SubConGetRs
             !AmountDue = !AmountDue + SubConDueRs!TotalDue
             !Total = !BFA + !AmountDue
             !CFA = !Total - !AmountPaid
             .Update
           End With 'SubConGetRs
         ElseIf SubConGetRs.RecordCount = 0 Then
           Set SubConGetRs = New ADODB.Recordset
           sqlString = "Select * from SubConPayments"
           SubConGetRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
           With SubConGetRs
            Set SubConPayRs = New ADODB.Recordset
            sqlString = "Select * From SubConpayments " & _
            "Where SubConNum=" & SubConRs!SubConNum
            SubConPayRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
            If SubConPayRs.RecordCount <> 0 Then
            SubConPayRs.MoveLast
            With SubConPayRs
            CFA = FormatCurrency(!CFA, 2)
            End With 'SubConPayRs
            End If 'SubConPayRs.recordcount
            .AddNew
            !SubConNum = SubConRs!SubConNum
            !Date = Format(TodayD, "dd/mm/yyyy")
            !ChequeNo = ""
            !BFA = CFA
            !AmountDue = SubConDueRs!TotalDue
            !AmountPaid = FormatCurrency(0, 2)
            !Total = !BFA + !AmountDue
            !CFA = !Total - !AmountPaid
            .Update
           End With 'SubConGetRs
         End If 'SubConGetRs.RecordCount
        End With 'SubConDueRs
       End If 'SubConDueRs.RecordCount
     SubConRs.MoveNext
     prgExport.Value = prgExport.Value + 1
     Next
   End With 'SubConRs
  End If 'SubCOnRs.RecordCount
  SubConRs.Close
End Sub
Private Sub AddNewConCon()
Dim ConNormalRs As ADODB.Recordset
Dim ConWShopRs As ADODB.Recordset
Dim NAmount As Currency
Dim WAmount As Currency
Dim TodayD As String
Dim ContractNo As String
Dim r As Integer
prgExport.Value = 0
TodayD = Date
Set ConConRS = New ADODB.Recordset
sqlString = "Select ContractNo from ContractDetails"
ConConRS.Open sqlString, DB, adOpenStatic, adLockOptimistic
If ConConRS.RecordCount <> 0 Then
  ConConRS.MoveFirst
  With ConConRS
   prgExport.Max = .RecordCount
   For r = 1 To ConConRS.RecordCount
     ContractNo = !ContractNo
     Set ConNormalRs = New ADODB.Recordset 'Get Total for Normal Work
     sqlString = "Select sum(Amount) as TotalNormal  from SubConWorkDetails where " & _
     "ContractNo='" & !ContractNo & "'" & _
     " and WorkType='Normal' and TransferDate is Null"
     ConNormalRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If ConNormalRs.RecordCount <> 0 Then
      With ConNormalRs
       If !TotalNormal >= 0 Then 'To Stop from getting error msg on value of Null
        NAmount = !TotalNormal
       Else
        NAmount = 0
       End If 'TotalNormal
      End With 'ConNormalRs.RecordCount
     End If 'if ConNormalRs.Recordcount
     
     Set ConWShopRs = New ADODB.Recordset 'Get Total for Workshop Work
     sqlString = "Select sum(Amount) as TotalWShop from SubConWorkDetails where " & _
     "ContractNo='" & !ContractNo & "'" & _
     " and WorkType='Workshop' and TransferDate is Null"
     ConWShopRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If ConWShopRs.RecordCount <> 0 Then
       With ConWShopRs
        If !TotalWShop >= 0 Then 'To Stop from getting error msg on value of Null
          WAmount = !TotalWShop
        Else
        WAmount = 0
        End If 'TotalWShop
       End With 'ConWShoprs.recordcount
     End If 'if ConWShopRs.RecordCount
    
    If NAmount + WAmount <> 0 Then
     Set ConExpRs = New ADODB.Recordset
     sqlString = "Select * From ContractExpenses where ContractNo='" & _
     !ContractNo & "' and Date=#" & Format(TodayD, "m/d/yyyy") & "#"
     ConExpRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If ConExpRs.RecordCount <> 0 Then
     With ConExpRs
     'Update Existing Record
     !laborSubContract = FormatCurrency(!laborSubContract + NAmount, 2)
     !LaborManuf = FormatCurrency(!LaborManuf + WAmount, 2)
     !Total = FormatCurrency(!Matrial + !LaborManuf + !laborSubContract + !Commission + !Overheads + !Total)
     .Update
     End With
     ElseIf ConExpRs.RecordCount = 0 Then
     'Add New Record
      Set ConExpRs = New ADODB.Recordset
      sqlString = "Select * from ContractExpenses"
      ConExpRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
      With ConExpRs
      .AddNew
      !ContractNo = ContractNo
      !Date = Format(TodayD, "dd/mm/yyyy")
      !Matrial = FormatCurrency(0, 2)
      !LaborManuf = FormatCurrency(WAmount, 2)
      !laborSubContract = FormatCurrency(NAmount, 2)
      !Commission = FormatCurrency(0, 2)
      !Overheads = FormatCurrency(0, 2)
      !other = FormatCurrency(0, 2)
      !Total = FormatCurrency(!Matrial + !LaborManuf + !laborSubContract + !Commission + !Overheads + !Total)
      .Update
      End With
      End If 'ConExpRs.recordcount
     End If 'NAmount + WAmount
   ConConRS.MoveNext
   prgExport.Value = prgExport.Value + 1
   Next 'For r
  End With 'ConConRs
End If 'ConConRs.Recordcount
End Sub
Private Sub TimeStamp()
Dim r As Integer
Dim TodayD As String
  TodayD = Date
  prgExport.Value = 0
  'Add Time Stamp to Exported Records
  Set SubWorkRs = New ADODB.Recordset
  sqlString = "Select * from SubconworkDetails where TransferDate is Null"
  SubWorkRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
  If SubWorkRs.RecordCount <> 0 Then
  prgExport.Max = SubWorkRs.RecordCount
  SubWorkRs.MoveFirst
   For r = 1 To SubWorkRs.RecordCount
     With SubWorkRs
     !TransferDate = Format(TodayD, "dd/mm/yyyy")
     .Update
     End With
     SubWorkRs.MoveNext
     prgExport.Value = prgExport.Value + 1
   Next
  End If 'Subworkrs.recordcount
End Sub

Private Sub Form_Load()
'SetRegion
End Sub
