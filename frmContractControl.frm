VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContractControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expences and Payments"
   ClientHeight    =   7305
   ClientLeft      =   390
   ClientTop       =   1170
   ClientWidth     =   11775
   Icon            =   "frmContractControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10320
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "Summary"
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   6840
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDeleteExp 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdateExp 
      Caption         =   "Update"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveExp 
      Caption         =   "Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtEditExp 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSMask.MaskEdBox mskEditExp 
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskEditPay 
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtEditPay 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeletePay 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdatePay 
      Caption         =   "Update"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdSavePay 
      Caption         =   "Insert"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPay 
      Height          =   1815
      Left            =   5880
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExp 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView lstContracts 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Payment Recieved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Expences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   11535
   End
End
Attribute VB_Name = "frmContractControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LConRs As ADODB.Recordset
Dim ConExpRs As ADODB.Recordset
Dim ConPayRs As ADODB.Recordset
Dim CheckRs As ADODB.Recordset
Dim Total As Currency
Dim sqlString As String

Private Sub DoList()
Dim ItmX As ListItem
lstContracts.ColumnHeaders.Add , , "Contract Num", Len("Contract Num") * 100
lstContracts.ColumnHeaders.Add , , "Client Name", Len("Client Name") * 100 + 1500
lstContracts.ColumnHeaders.Add , , "Amount", lstContracts.Width - (Len("Contract Num") * 100 + Len("Client Name") * 100 + 1500) - 50 'Len("Amount") * 100 + 800
    Set LConRs = New ADODB.Recordset
    LConRs.Open "SELECT * FROM ContractDetails", DB, adOpenStatic, adLockOptimistic
    If Not LConRs.BOF Then LConRs.MoveFirst
    lstContracts.ListItems.Clear
    Do While Not LConRs.EOF
        Set ItmX = lstContracts.ListItems.Add(, , LConRs!ContractNo)
         ItmX.SubItems(1) = LConRs!Clientname
         ItmX.SubItems(2) = FormatCurrency(LConRs!Amount, 2)
        LConRs.MoveNext
    Loop
    If Not LConRs.EOF Then LConRs.MoveFirst
    lstContracts.Refresh
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdDeleteExp_Click()
Dim ControlDel As ADODB.Recordset
Dim ContNum As String
Dim ConDate As String
If ConExpRs.RecordCount <> 0 Then
    ConExpRs.MoveFirst
    If Not ConExpRs.EOF Or ConExpRs.BOF Then
        If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
         Set ControlDel = New ADODB.Recordset
         grdExp.Row = grdExp.Row
         ContNum = lstContracts.SelectedItem
         grdExp.Col = 1: ConDate = grdExp.Text
         sqlString = "Select * from ContractExpenses where ContractNo='" & _
         Trim(ContNum) & "' and Date=#" & Format(ConDate, "m/d/yyyy") & "#"
         ControlDel.Open sqlString, DB, adOpenStatic, adLockOptimistic
         If ControlDel.RecordCount <> 0 Then
          ControlDel.Delete
         End If
         ClearGridExp
         CreateGridExp
         GetExpData
         End If
        End If
 End If
End Sub

Private Sub cmdDeletePay_Click()
Dim ControlDel As ADODB.Recordset
Dim ContNum As String
Dim ConDate As Date
If ConPayRs.RecordCount <> 0 Then
    ConPayRs.MoveFirst
    If Not ConPayRs.EOF Or ConPayRs.BOF Then
        If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
         Set ControlDel = New ADODB.Recordset
         grdPay.Row = grdPay.Row
         ContNum = lstContracts.SelectedItem
         grdPay.Col = 1: ConDate = grdPay.Text
         sqlString = "Select * from PaymentsRecieved where ContractNo='" & _
         Trim(ContNum) & "' and Date=#" & Format(ConDate, "m/d/yyyy") & "#"
         ControlDel.Open sqlString, DB, adOpenStatic, adLockOptimistic
         If ControlDel.RecordCount <> 0 Then
          ControlDel.Delete
         End If
         ClearGridPay
         CreateGridPay
         GetPayData
         End If
        End If
End If
End Sub

Private Sub cmdSaveExp_Click()
Dim IDate
Dim r As Integer
If cmdSaveExp.Caption = "Insert" Then
cmdSaveExp.Caption = "Save"
cmdUpdateExp.Enabled = False
cmdDeleteExp.Enabled = False
ClearGridExp
CreateGridExp
grdExp.SetFocus
grdExp.Col = 1
grdExp.Rows = grdExp.Rows
grdExp.Row = 1
grdExp_EnterCell
ElseIf cmdSaveExp.Caption = "Save" Then
cmdSaveExp.Caption = "Insert"
cmdUpdateExp.Enabled = True
cmdDeleteExp.Enabled = True
     Set ConExpRs = New ADODB.Recordset
     ConExpRs.Open "Select * from ContractExpenses", DB, adOpenStatic, adLockOptimistic
     For r = 1 To grdExp.Rows - 2
        grdExp.Row = r
        With ConExpRs
            .AddNew
            !ContractNo = lstContracts.SelectedItem
            grdExp.Col = 1: IDate = grdExp.Text
            grdExp.Col = 1: !Date = CDate(IDate)
            grdExp.Col = 2: !Item = grdExp.Text
            grdExp.Col = 3: !Matrial = grdExp.Text
            grdExp.Col = 4: !LaborManuf = grdExp.Text
            grdExp.Col = 5: !laborSubContract = grdExp.Text
            grdExp.Col = 6: !Commission = grdExp.Text
            grdExp.Col = 7: !Overheads = grdExp.Text
            grdExp.Col = 8: !other = grdExp.Text
            grdExp.Col = 9: !Total = grdExp.Text
            .Update
        End With
     Next
     ClearGridExp
     CreateGridExp
     GetExpData
End If
End Sub

Private Sub cmdSavePay_Click()
Dim IDate
Dim r As Integer
If cmdSavePay.Caption = "Insert" Then
  cmdSavePay.Caption = "Save"
  cmdUpdatePay.Enabled = False
  cmdDeletePay.Enabled = False
  ClearGridPay
  CreateGridPay
  grdPay.SetFocus
  grdPay.Col = 1
  grdPay.Rows = grdPay.Rows
  grdPay.Row = 1
  grdPay_EnterCell
ElseIf cmdSavePay.Caption = "Save" Then
  cmdSavePay.Caption = "Insert"
  cmdUpdatePay.Enabled = True
  cmdDeletePay.Enabled = True
     Set ConPayRs = New ADODB.Recordset
     ConPayRs.Open "Select * from PaymentsRecieved", DB, adOpenStatic, adLockOptimistic
     For r = 1 To grdPay.Rows - 2
        grdPay.Row = r
        With ConPayRs
            .AddNew
            !ContractNo = lstContracts.SelectedItem
            '!subconnum = lstSubCon.SelectedItem.SubItems(1)
            grdPay.Col = 1: IDate = grdPay.Text
            grdPay.Col = 1: !Date = CDate(IDate)
            grdPay.Col = 2: !PaymentDesc = grdPay.Text
            grdPay.Col = 3: !AmountPaid = grdPay.Text
            .Update
        End With
     Next
     ClearGridPay
     CreateGridPay
     GetPayData
End If
End Sub

Private Sub cmdSummary_Click()
frmConControlSummary.Show
End Sub

Private Sub cmdUpdateExp_Click()
Dim IDate
Dim r As Integer
Dim ConNum As String
     For r = 1 To grdExp.Rows - 1
     grdExp.Row = r
     ConNum = lstContracts.SelectedItem
     grdExp.Col = 1: IDate = grdExp.Text
     Set ConExpRs = New ADODB.Recordset
     sqlString = "Select * from ContractExpenses where ContractNo='" & Trim(ConNum) _
     & "' and Date=#" & Format(IDate, "m/d/yyyy") & "#"
     ConExpRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If ConExpRs.RecordCount <> 0 Then
        With ConExpRs
            '!subconnum = lstSubCon.SelectedItem.SubItems(1)
            '!Contractno = SubNum
            '!Date = CDate(IDate)
            'grdWork.Col = 3: !WorkDescription = grdWork.Text
            'grdWork.Col = 4: !Worktype = grdWork.Text
            'grdWork.Col = 5: !Amount = grdWork.Text
            !ContractNo = lstContracts.SelectedItem
            !Date = CDate(IDate)
            grdExp.Col = 2: !Item = grdExp.Text
            grdExp.Col = 3: !Matrial = grdExp.Text
            grdExp.Col = 4: !LaborManuf = grdExp.Text
            grdExp.Col = 5: !laborSubContract = grdExp.Text
            grdExp.Col = 6: !Commission = grdExp.Text
            grdExp.Col = 7: !Overheads = grdExp.Text
            grdExp.Col = 8: !other = grdExp.Text
            grdExp.Col = 9: !Total = grdExp.Text
            .Update
        End With
     End If
     Next
     ClearGridPay
     CreateGridPay
     GetPayData

End Sub

Private Sub cmdUpdatePay_Click()
Dim IDate
Dim r As Integer
Dim ConNum As String
     For r = 1 To grdPay.Rows - 1
     grdPay.Row = r
     ConNum = lstContracts.SelectedItem
     grdPay.Col = 1: IDate = grdPay.Text
     Set ConPayRs = New ADODB.Recordset
     sqlString = "Select * from PaymentsRecieved where ContractNo='" & Trim(ConNum) _
     & "' and Date=#" & Format(IDate, "m/d/yyyy") & "#"
     ConPayRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If ConPayRs.RecordCount <> 0 Then
        With ConPayRs
            !ContractNo = lstContracts.SelectedItem
            !Date = CDate(IDate)
            grdPay.Col = 2: !PaymentDesc = grdPay.Text
            grdPay.Col = 3: !AmountPaid = grdPay.Text
            .Update
        End With
     End If
     Next
     ClearGridPay
     CreateGridPay
     GetPayData
End Sub

Private Sub Form_Load()
CreateGridExp
CreateGridPay
DoList
End Sub

Private Sub CreateGridExp()
grdExp.Cols = 10
grdExp.Rows = 2
grdExp.Row = 0
grdExp.Col = 1: grdExp.Text = "Date"
grdExp.Col = 2: grdExp.Text = "Item"
grdExp.Col = 3: grdExp.Text = "Material"
grdExp.Col = 4: grdExp.Text = "Labor Manuf"
grdExp.Col = 5: grdExp.Text = "Labor Sub Con"
grdExp.Col = 6: grdExp.Text = "Commission"
grdExp.Col = 7: grdExp.Text = "Overheads"
grdExp.Col = 8: grdExp.Text = "Other"
grdExp.Col = 9: grdExp.Text = "Total"
grdExp.ColWidth(0) = 300
grdExp.ColWidth(1) = Len("12/12/2001") * 95
grdExp.ColWidth(2) = Len("Item Description") * 80
grdExp.ColWidth(3) = Len("Labor Sub Con") * 95
grdExp.ColWidth(4) = Len("Labor Sub Con") * 95
grdExp.ColWidth(5) = Len("Labor Sub Con") * 95
grdExp.ColWidth(6) = Len("Labor Sub Con") * 95
grdExp.ColWidth(7) = Len("Labor Sub Con") * 95
grdExp.ColWidth(8) = Len("Labor Sub Con") * 95
grdExp.ColWidth(9) = Len("Labor Sub Con") * 95
End Sub

Private Sub CreateGridPay()
grdPay.Cols = 4
grdPay.Rows = 2
grdPay.Row = 0
grdPay.Col = 1: grdPay.Text = "Date"
grdPay.Col = 2: grdPay.Text = "Description"
grdPay.Col = 3: grdPay.Text = "Amount Paid"
grdPay.ColWidth(0) = 300
grdPay.ColWidth(1) = Len("12/12/2001") * 100
grdPay.ColWidth(2) = Len("Description") * 100 + 1650
grdPay.ColWidth(3) = Len("Amount Paid") * 100 + 250
End Sub

Private Sub grdExp_EnterCell()
    '// when click on cell
    Select Case grdExp.Col
        Case 1 'Date
         If cmdSaveExp.Caption = "Save" Then
         With mskEditExp
             .Move grdExp.CellLeft + grdExp.Left, _
             grdExp.CellTop + grdExp.Top, grdExp.CellWidth - 25, _
             grdExp.CellHeight - 25
             .SelText = grdExp.Text
             If Len(.Text) > 0 Then
               .SelStart = 0
               .SelLength = Len(.Text)
             End If
             .Visible = True
             .ZOrder 0
             .SetFocus
         End With
         ElseIf cmdSaveExp.Caption = "Insert" Then
         grdExp.Col = 2
         grdExp_EnterCell
         End If
        Case 2, 3, 4, 5, 6, 7, 8, 9
         With txtEditExp
             .Move grdExp.CellLeft + grdExp.Left, _
             grdExp.CellTop + grdExp.Top, grdExp.CellWidth - 25, _
             grdExp.CellHeight - 25
             .Text = grdExp.Text
             If Len(.Text) > 0 Then
               .SelStart = 0
               .SelLength = Len(.Text)
             End If
             .Visible = True
             .ZOrder 0
             .SetFocus
         End With
    End Select

End Sub

Private Sub grdPay_EnterCell()
    '// when click on cell
    Select Case grdPay.Col
        Case 1 'Date
         If cmdSavePay.Caption = "Save" Then
         With mskEditPay
             .Move grdPay.CellLeft + grdPay.Left, _
             grdPay.CellTop + grdPay.Top, grdPay.CellWidth - 25, _
             grdPay.CellHeight - 25
             .SelText = grdPay.Text
             If Len(.Text) > 0 Then
               .SelStart = 0
               .SelLength = Len(.Text)
             End If
             .Visible = True
             .ZOrder 0
             .SetFocus
         End With
         ElseIf cmdSavePay.Caption = "Insert" Then
         grdPay.Col = 2
         grdPay_EnterCell
         End If
        Case 2, 3 'Payment Description, AmountPaid
         With txtEditPay
             .Move grdPay.CellLeft + grdPay.Left, _
             grdPay.CellTop + grdPay.Top, grdPay.CellWidth - 25, _
             grdPay.CellHeight - 25
             .Text = grdPay.Text
             If Len(.Text) > 0 Then
               .SelStart = 0
               .SelLength = Len(.Text)
             End If
             .Visible = True
             .ZOrder 0
             .SetFocus
         End With
    End Select
End Sub

Private Sub lstContracts_Click()
prgLoad.Value = 0
cmdSaveExp.Caption = "Insert"
cmdSavePay.Caption = "Insert"
cmdUpdateExp.Enabled = True
cmdDeleteExp.Enabled = True
cmdUpdatePay.Enabled = True
cmdDeletePay.Enabled = True
txtEditPay.Visible = False
mskEditPay.Visible = False
txtEditExp.Visible = False
mskEditExp.Visible = False
FromListUpdateGrid
End Sub

Private Sub lstContracts_KeyUp(KeyCode As Integer, Shift As Integer)
prgLoad.Value = 0
cmdSaveExp.Caption = "Insert"
cmdSavePay.Caption = "Insert"
cmdUpdateExp.Enabled = True
cmdDeleteExp.Enabled = True
cmdUpdatePay.Enabled = True
cmdDeletePay.Enabled = True
txtEditPay.Visible = False
mskEditPay.Visible = False
txtEditExp.Visible = False
mskEditExp.Visible = False
FromListUpdateGrid
End Sub

Private Sub mskEditExp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With mskEditExp
                    .Text = Empty
                    .Visible = False
                End With
                grdExp.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With mskEditExp
                    If Not .Text = Empty Then
                        If Not IsDate(.Text) Then
                          MsgBox "Not a valid date"
                          .SetFocus
                          Exit Sub
                        End If
                        Set CheckRs = New ADODB.Recordset
                        sqlString = "Select * from ContractExpenses where ContractNo='" & _
                        Trim(lstContracts.SelectedItem) & "'" & _
                        " and Date=#" & Format(.Text, "m/d/yyyy") & "#"
                        CheckRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
                         If CheckRs.RecordCount <> 0 Then
                          MsgBox "Cannot have duplicate date entries for an expence"
                          .SetFocus
                          Exit Sub
                         End If
                        
                        grdExp.Text = .Text
                    End If
                 
                    .Visible = False
                    .SelText = Empty
                End With
                Select Case grdExp.Col
                    Case 1
                        grdExp.Col = 2
                        grdExp_EnterCell
                End Select
    End Select
End Sub

Private Sub mskEditPay_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With mskEditPay
                    .Text = Empty
                    .Visible = False
                End With
                grdPay.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With mskEditPay
                    If Not .Text = Empty Then
                        If Not IsDate(.Text) Then
                          MsgBox "Not a valid date"
                          .SetFocus
                          Exit Sub
                        End If
                        Set CheckRs = New ADODB.Recordset
                        sqlString = "Select * from PaymentsRecieved where ContractNo='" & _
                        Trim(lstContracts.SelectedItem) & "'" & _
                        " and Date=#" & Format(.Text, "m/d/yyyy") & "#"
                        CheckRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
                         If CheckRs.RecordCount <> 0 Then
                          MsgBox "Only one payment entrie is allowed per date"
                          .SetFocus
                          Exit Sub
                         End If

                        grdPay.Text = .Text
                    End If
                 
                    .Visible = False
                    .SelText = Empty
                End With
                Select Case grdPay.Col
                    Case 1
                        grdPay.Col = 2
                        grdPay_EnterCell
               End Select
    End Select
End Sub

Private Sub txtEditExp_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With txtEditExp
                    .Text = Empty
                    .Visible = False
                End With
                grdExp.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With txtEditExp
                    If Not .Text = Empty Then
                        grdExp.Text = .Text
                    ElseIf .Text = "" And grdExp.Col = 3 Or grdExp.Col = 4 Or _
                    grdExp.Col = 5 Or grdExp.Col = 6 Or grdExp.Col = 7 Or grdExp.Col = 8 Then
                       grdExp.Text = 0
                    End If
                    .Visible = False
                    .SelText = Empty
                End With
                Select Case grdExp.Col
                    Case 2
                        grdExp.Col = 3
                        grdExp_EnterCell
                    Case 3
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 4
                        grdExp_EnterCell
                    Case 4
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 5
                        grdExp_EnterCell
                    Case 5
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 6
                        grdExp_EnterCell
                    Case 6
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 7
                        grdExp_EnterCell
                    Case 7
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 8
                        grdExp_EnterCell
                    Case 8
                        grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        grdExp.Col = 9
                        'grdExp_EnterCell
                        Totals
                    'Case 9
                        grdExp.Col = 9: grdExp.Text = Total
                        If grdExp.Col = 9 And grdExp.Text <> "" And cmdSaveExp.Caption = "Save" Then
                          grdExp.Text = FormatCurrency(grdExp.Text, 2)
                          grdExp.Rows = grdExp.Rows + 1
                          grdExp.Row = grdExp.Row + 1
                          grdExp.Col = 1
                          grdExp_EnterCell
                        Else
                          grdExp.Text = FormatCurrency(grdExp.Text, 2)
                        End If
                End Select
    End Select
End Sub

Private Sub txtEditExp_KeyPress(KeyAscii As Integer)
Select Case grdExp.Col
    Case 3, 4, 5, 6, 7, 8
     KeyAscii = ValidateInput(KeyAscii, Currency_Input)
End Select
End Sub

Private Sub txtEditPay_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With txtEditPay
                    .Text = Empty
                    .Visible = False
                End With
                grdPay.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With txtEditPay
                    If Not .Text = Empty Then
                        grdPay.Text = .Text
                    End If
                    .Visible = False
                    .SelText = Empty
                End With
                Select Case grdPay.Col
                    Case 2
                        grdPay.Col = 3
                        grdPay_EnterCell
                    Case 3
                        If grdPay.Col = 3 And grdPay.Text <> "" And cmdSavePay.Caption = "Save" Then
                            grdPay.Text = FormatCurrency(grdPay.Text, 2)
                            grdPay.Rows = grdPay.Rows + 1
                            grdPay.Row = grdPay.Row + 1
                            grdPay.Col = 1
                            grdPay_EnterCell
                         Else
                         grdPay.Text = FormatCurrency(grdPay.Text, 2)
                         End If
                End Select
    End Select
    'Case 5
    'If grdWork.Col = 5 And grdWork.Text <> "" And optInsert.Value = True Then
    'grdWork.Text = FormatCurrency(grdWork.Text, 2)
    'grdWork.Rows = grdWork.Rows + 1
    'grdWork.Row = grdWork.Row + 1
    'grdWork.Col = 1
    'grdWork_EnterCell
    'elseIf grdWork.Row <> grdWork.Rows - 1 Then
    'grdWork.Text = FormatCurrency(grdWork.Text, 2)
    'grdWork.Row = grdWork.Row + 1
    'grdWork.Col = 1
    'grdWork_EnterCell
    'Else
    'grdWork.Text = FormatCurrency(grdWork.Text, 2)
    'End If
End Sub

Private Sub txtEditPay_KeyPress(KeyAscii As Integer)
If grdPay.Col = 3 Then
     KeyAscii = ValidateInput(KeyAscii, Currency_Input)
End If
End Sub

Private Sub Totals()
Dim r As Integer
Total = 0
For r = 3 To 8
grdExp.Col = r
If grdExp.Text <> "" Then
Total = Total + grdExp.Text
Else
Total = Total + 0
End If
'grdExp.Col = 4: Total = Total + grdExp.Text
'grdExp.Col = 5: Total = Total + grdExp.Text
'grdExp.Col = 6: Total = Total + grdExp.Text
'grdExp.Col = 7: Total = Total + grdExp.Text
'grdExp.Col = 8: Total = Total + grdExp.Text
Next
End Sub

Private Sub FromListUpdateGrid()
 
 GetPayData
 GetExpData
End Sub
Private Sub GetPayData()
Dim r As Integer
TotalPay = 0
prgLoad.Value = 0
Set ConPayRs = New ADODB.Recordset
 If Not lstContracts.SelectedItem.Text = Empty Then
    sqlString = Trim(lstContracts.SelectedItem.Text)
    ConPayRs.Open "Select * from PaymentsRecieved where Contractno='" & sqlString & "'", DB, adOpenStatic, adLockOptimistic
     grdPay.Clear
     CreateGridPay
   If ConPayRs.RecordCount <> 0 Then
    prgLoad.Max = ConPayRs.RecordCount
   Else
    prgLoad.Max = 1
   End If
   grdPay.Row = grdPay.Row
   grdPay.Col = 1
    If ConPayRs.RecordCount <> 0 Then
      For r = 1 To ConPayRs.RecordCount
       With ConPayRs
         grdPay.Rows = grdPay.Rows + 1
         prgLoad.Value = r
         grdPay.Col = 1: grdPay.Text = Format(!Date, "dd/mm/yyyy")
         grdPay.Col = 2: grdPay.Text = !PaymentDesc
         grdPay.Col = 3: grdPay.Text = FormatCurrency(!AmountPaid, 2)
         TotalPay = TotalPay + !AmountPaid
         grdPay.Row = grdPay.Row + 1
         grdPay.Col = 1
         .MoveNext
       End With
      Next
    End If
   If grdPay.Rows > 2 Then
    grdPay.Rows = grdPay.Rows - 1
   End If
 End If
End Sub
Private Sub GetExpData()
Dim r As Integer
TotalExp = 0
prgLoad.Value = 0
Set ConExpRs = New ADODB.Recordset
 If Not lstContracts.SelectedItem.Text = Empty Then
    sqlString = Trim(lstContracts.SelectedItem.Text)
    ConExpRs.Open "Select * from ContractExpenses where Contractno='" & sqlString & "'", DB, adOpenStatic, adLockOptimistic
     grdExp.Clear
     CreateGridExp
   If ConExpRs.RecordCount <> 0 Then
    prgLoad.Max = ConExpRs.RecordCount
   Else
    prgLoad.Max = 1
   End If
   grdExp.Row = grdExp.Row
   grdExp.Col = 1
    If ConExpRs.RecordCount <> 0 Then
      For r = 1 To ConExpRs.RecordCount
       With ConExpRs
         grdExp.Rows = grdExp.Rows + 1
         prgLoad.Value = r
         grdExp.Col = 1: grdExp.Text = Format(!Date, "dd/mm/yyyy")
         
         grdExp.Col = 2
         If !Item <> "" Then
         grdExp.Text = !Item
         Else
         grdExp.Text = ""
         End If
         grdExp.Col = 3: grdExp.Text = FormatCurrency(!Matrial, 2)
         grdExp.Col = 4: grdExp.Text = FormatCurrency(!LaborManuf, 2)
         grdExp.Col = 5: grdExp.Text = FormatCurrency(!laborSubContract, 2)
         grdExp.Col = 6: grdExp.Text = FormatCurrency(!Commission, 2)
         grdExp.Col = 7: grdExp.Text = FormatCurrency(!Overheads, 2)
         grdExp.Col = 8: grdExp.Text = FormatCurrency(!other, 2)
         grdExp.Col = 9: grdExp.Text = FormatCurrency(!Total, 2)
         TotalExp = TotalExp + !Total
         grdExp.Row = grdExp.Row + 1
         grdExp.Col = 1
         .MoveNext
       End With
      Next
    End If
   If grdExp.Rows > 2 Then
    grdExp.Rows = grdExp.Rows - 1
   End If
 End If
End Sub

Private Sub ClearGridPay()
grdPay.Clear
txtEditPay.Visible = False
mskEditPay.Visible = False
txtEditExp.Visible = False
mskEditExp.Visible = False
End Sub

Private Sub ClearGridExp()
grdExp.Clear
txtEditPay.Visible = False
mskEditPay.Visible = False
txtEditExp.Visible = False
mskEditExp.Visible = False
End Sub
