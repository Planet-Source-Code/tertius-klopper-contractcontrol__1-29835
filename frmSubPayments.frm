VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontractor Payments"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   Icon            =   "frmSubPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   135
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Insert"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPay 
      Height          =   4695
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8281
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.ListView lstSubCon 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9551
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
End
Attribute VB_Name = "frmSubPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CheckRs As ADODB.Recordset
Dim LSubRs As ADODB.Recordset
Dim SubPayRs As ADODB.Recordset
Dim sqlString As String
Dim CheckToday As Boolean

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim SubConDel As ADODB.Recordset
Dim ContNum As String
Dim ConDate As String

If SubPayRs.RecordCount <> 0 Then
         SubPayRs.MoveFirst
    If Not SubPayRs.EOF Or SubPayRs.BOF Then
        If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
         Set SubConDel = New ADODB.Recordset
         grdPay.Row = grdPay.Row
         grdPay.Col = 1: ConDate = grdPay.Text
         sqlString = "Select * from subconpayments where SubConNum=" & _
         Trim(lstSubCon.SelectedItem.SubItems(1)) & _
         "and Date=#" & Format(ConDate, "m/d/yyyy") & "#"
         SubConDel.Open sqlString, DB, adOpenStatic, adLockOptimistic
         If SubConDel.RecordCount <> 0 Then
          SubConDel.Delete
         End If
         ClearGrid
         FromListUpdateRecord
         End If
        End If
End If
End Sub

Private Sub cmdPrint_Click()
frmSubConReportSelect.optPaymentrpt.Value = True
frmSubConReportSelect.Show
End Sub

Private Sub cmdSave_Click()
Dim TodayD As String

Dim IDate
Dim r As Integer

If cmdSave.Caption = "Insert" Then
 CheckToday = False
 cmdSave.Caption = "Save"
 cmdUpdate.Enabled = False
 cmdDelete.Enabled = False
 ClearGrid
 CreateGrid
 grdPay.Row = 1
 TodayD = Date
 Set SubPayRs = New ADODB.Recordset
  sqlString = "Select * from SubConPayments where SubConNum=" & _
  Trim(lstSubCon.SelectedItem.SubItems(1)) & _
  " and Date=#" & Format(TodayD, "m/d/yyyy") & "#"
  SubPayRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
   If SubPayRs.RecordCount <> 0 Then
    CheckToday = True
       With SubPayRs
         grdPay.Col = 1: grdPay.Text = Format(!Date, "dd/mm/yyyy")
         grdPay.Col = 2: grdPay.Text = !ChequeNo
         grdPay.Col = 3: grdPay.Text = FormatCurrency(!BFA, 2)
         grdPay.Col = 4: grdPay.Text = FormatCurrency(!AmountDue, 2)
         grdPay.Col = 5: grdPay.Text = FormatCurrency(!Total, 2)
         grdPay.Col = 6: grdPay.Text = FormatCurrency(!AmountPaid, 2)
         grdPay.Col = 7: grdPay.Text = FormatCurrency(!CFA, 2)
       End With
   ElseIf SubPayRs.RecordCount = 0 Then
   CheckToday = False
   Set SubPayRs = New ADODB.Recordset
   sqlString = "Select * from SubConPayments where SubConNum=" & _
   Trim(lstSubCon.SelectedItem.SubItems(1))
   SubPayRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
   If SubPayRs.RecordCount <> 0 Then
   SubPayRs.MoveLast
    With SubPayRs
      grdPay.Col = 1: grdPay.Text = Format(TodayD, "dd/mm/yyyy")
      grdPay.Col = 3: grdPay.Text = FormatCurrency(!CFA, 2)
      grdPay.Col = 2
      grdPay_EnterCell
    End With
   Else
    grdPay.Col = 1
    grdPay_EnterCell
   End If
   End If
ElseIf cmdSave.Caption = "Save" And CheckToday <> True Then
cmdSave.Caption = "Insert"
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
     Set SubPayRs = New ADODB.Recordset
     SubPayRs.Open "Select * from SubconPayments", DB, adOpenStatic, adLockOptimistic
     For r = 1 To grdPay.Rows - 1
        grdPay.Row = r
        With SubPayRs
            .AddNew
            !SubConNum = lstSubCon.SelectedItem.SubItems(1)
            grdPay.Col = 1: IDate = grdPay.Text
            grdPay.Col = 1: !Date = IDate
            grdPay.Col = 2: !ChequeNo = grdPay.Text
            grdPay.Col = 3: !BFA = grdPay.Text
            grdPay.Col = 4: !AmountDue = grdPay.Text
            grdPay.Col = 5: !Total = grdPay.Text
            grdPay.Col = 6: !AmountPaid = grdPay.Text
            grdPay.Col = 7: !CFA = grdPay.Text
            .Update
        End With
     Next
     ClearGrid
     CreateGrid
     FromListUpdateRecord
     
ElseIf cmdSave.Caption = "Save" And CheckToday = True Then
cmdSave.Caption = "Insert"
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
'     For r = 1 To grdPay.Rows - 1
        grdPay.Row = grdPay.Row
        With SubPayRs
            !SubConNum = lstSubCon.SelectedItem.SubItems(1)
            grdPay.Col = 1: IDate = grdPay.Text
            grdPay.Col = 1: !Date = IDate
            grdPay.Col = 2: !ChequeNo = grdPay.Text
            grdPay.Col = 3: !BFA = grdPay.Text
            grdPay.Col = 4: !AmountDue = grdPay.Text
            grdPay.Col = 5: !Total = grdPay.Text
            grdPay.Col = 6: !AmountPaid = grdPay.Text
            grdPay.Col = 7: !CFA = grdPay.Text
            .Update
        End With
'     Next
     ClearGrid
     CreateGrid
     FromListUpdateRecord
End If
End Sub

Private Sub cmdUpdate_Click()
Dim IDate

 grdPay.Row = grdPay.Rows - 1
 grdPay.Col = 1: IDate = grdPay.Text
 Set SubPayRs = New ADODB.Recordset
 sqlString = "Select * from SubConPayments where SubConNum=" & _
 Trim(lstSubCon.SelectedItem.SubItems(1)) & _
 " and Date=#" & Format(IDate, "m/d/yyyy") & "#"
 SubPayRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
 If SubPayRs.RecordCount <> 0 Then
   With SubPayRs
   !SubConNum = lstSubCon.SelectedItem.SubItems(1)
   !Date = CDate(IDate)
   grdPay.Col = 2: !ChequeNo = grdPay.Text
   grdPay.Col = 3: !BFA = grdPay.Text
   grdPay.Col = 4: !AmountDue = grdPay.Text
   grdPay.Col = 5: !Total = grdPay.Text
   grdPay.Col = 6: !AmountPaid = grdPay.Text
   grdPay.Col = 7: !CFA = grdPay.Text
   .Update
 End With
 ClearGrid
 CreateGrid
 FromListUpdateRecord
End If
End Sub

Private Sub Form_Load()
CreateGrid
DoList
End Sub
Private Sub DoList()
Dim ItmX As ListItem
lstSubCon.ColumnHeaders.Add , , "First Name, Last Name", lstSubCon.Width - 100
lstSubCon.ColumnHeaders.Add , , "SubConNum", 5
Set LSubRs = New ADODB.Recordset
    LSubRs.Open "SELECT * FROM SubconDetails", DB, adOpenStatic, adLockOptimistic
    If Not LSubRs.BOF Then LSubRs.MoveFirst
    lstSubCon.ListItems.Clear
    Do While Not LSubRs.EOF
        Set ItmX = lstSubCon.ListItems.Add(, , LSubRs!subconfname & ", " & LSubRs!subconLname)
        ItmX.SubItems(1) = LSubRs!SubConNum
        LSubRs.MoveNext
    Loop
    If Not LSubRs.EOF Then LSubRs.MoveFirst
    lstSubCon.Refresh
End Sub

Private Sub grdPay_EnterCell()
    Select Case grdPay.Col
        Case 1
            If cmdSave.Caption = "Save" And grdPay.Row = grdPay.Rows - 1 Then
            With mskEdit
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
           ElseIf cmdSave.Caption = "Insert" Then
             grdPay.Col = 2
             grdPay_EnterCell
           End If
        Case 2, 3, 4, 5, 6, 7
             If grdPay.Row = grdPay.Rows - 1 Then
             With txtEdit
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
            End If
    End Select

End Sub

Private Sub lstSubCon_Click()
txtEdit.Visible = False
mskEdit.Visible = False
cmdSave.Caption = "Insert"
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
FromListUpdateRecord
End Sub

Private Sub lstSubCon_KeyUp(KeyCode As Integer, Shift As Integer)
txtEdit.Visible = False
mskEdit.Visible = False
cmdSave.Caption = "Insert"
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
FromListUpdateRecord
End Sub
Private Sub FromListUpdateRecord()
Dim r As Integer
 Set SubPayRs = New ADODB.Recordset
 If Not lstSubCon.SelectedItem.Text = Empty Then
   sqlString = Trim(lstSubCon.SelectedItem.SubItems(1))
   SubPayRs.Open "Select * from SubConPayments where SubconNum=" & sqlString, DB, adOpenStatic, adLockOptimistic
   If SubPayRs.RecordCount <> 0 Then
    prgLoad.Max = SubPayRs.RecordCount
   Else
    prgLoad.Max = 1
   End If
   grdPay.Clear
   CreateGrid
   grdPay.Row = 1
   grdPay.Col = 1
    If SubPayRs.RecordCount <> 0 Then
      For r = 1 To SubPayRs.RecordCount
       With SubPayRs
         grdPay.Rows = grdPay.Rows + 1
         prgLoad.Value = r
         grdPay.Col = 1: grdPay.Text = Format(!Date, "dd/mm/yyyy")
         grdPay.Col = 2: grdPay.Text = !ChequeNo
         grdPay.Col = 3: grdPay.Text = FormatCurrency(!BFA, 2)
         grdPay.Col = 4: grdPay.Text = FormatCurrency(!AmountDue, 2)
         grdPay.Col = 5: grdPay.Text = FormatCurrency(!Total, 2)
         grdPay.Col = 6: grdPay.Text = FormatCurrency(!AmountPaid, 2)
         grdPay.Col = 7: grdPay.Text = FormatCurrency(!CFA, 2)
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

Private Sub CreateGrid()
grdPay.Cols = 8
grdPay.Rows = 2
grdPay.Row = 0
grdPay.Col = 1: grdPay.Text = "Date"
grdPay.Col = 2: grdPay.Text = "ChequeNo"
grdPay.Col = 3: grdPay.Text = "BFA"
grdPay.Col = 4: grdPay.Text = "Amount Due"
grdPay.Col = 5: grdPay.Text = "Total Amount Due"
grdPay.Col = 6: grdPay.Text = "Amount Paid"
'grdPay.Col = 7: grdPay.Text = "LBF"
'grdPay.Col = 8: grdPay.Text = "LPB"
grdPay.Col = 7: grdPay.Text = "CFA"
grdPay.ColWidth(0) = 300
grdPay.ColWidth(1) = Len("12/12/2001") * 95
grdPay.ColWidth(2) = Len("ChequeNo") * 95 + 250
grdPay.ColWidth(3) = Len("BFA") * 95 + 1000
grdPay.ColWidth(4) = Len("Amount Due") * 95 + 300
grdPay.ColWidth(5) = Len("Total Amount Due") * 95
grdPay.ColWidth(6) = Len("Amount Paid") * 95 + 200
'grdPay.ColWidth(7) = Len("LBF") * 95 + 1000
'grdPay.ColWidth(8) = Len("LPB") * 95 + 1000
grdPay.ColWidth(7) = Len("CFA") * 95 + 1000
End Sub

Private Sub mskEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CheckRs As ADODB.Recordset
Dim ConNum As String
    Select Case KeyCode
            Case vbKeyEscape
                With mskEdit
                    .Text = Empty
                    .SelText = Empty
                    .Visible = False
                End With
                grdPay.SetFocus
            Case vbKeyReturn
                With mskEdit
                    If Not .Text = Empty Then
                        If Not IsDate(.Text) Then
                          MsgBox "Not a valid date"
                          .SetFocus
                          Exit Sub
                        Else
                        Set CheckRs = New ADODB.Recordset
                        sqlString = "Select * from SubConPayments where SubConNum=" & _
                        Trim(lstSubCon.SelectedItem.SubItems(1)) & _
                        " and Date=#" & Format(.Text, "m/d/yyyy") & "#"
                        CheckRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
                         If CheckRs.RecordCount <> 0 Then
                          MsgBox "Only on payment entrie is allowed per date"
                          .SetFocus
                          Exit Sub
                         End If
                        grdPay.Text = .Text
                        End If
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

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ConNum As ADODB.Recordset
    Dim BFA As Currency
    Dim AmountDue As Currency
    Dim TotalDue As Currency
    Dim AmountPaid As Currency
    Dim CFA As Currency
    Select Case KeyCode
            Case vbKeyEscape
                With txtEdit
                    .Text = Empty
                    .Visible = False
                End With
                grdPay.SetFocus
          Case vbKeyReturn
                With txtEdit
                    If Not .Text = Empty Then
                      grdPay.Text = .Text
                      Select Case grdPay.Col
                       Case 3, 4, 5, 6, 7, 8
                        grdPay.Text = FormatCurrency(grdPay.Text, 2)
                      End Select
                    ElseIf .Text = "" Then
                    Select Case grdPay.Col
                      Case 3, 4, 5, 6, 7, 8
                        grdPay.Text = 0
                        grdPay.Text = FormatCurrency(grdPay.Text, 2)
                    End Select
                     End If
                    .Visible = False
                    .Text = Empty
                End With
                Select Case grdPay.Col
                    Case 2
                        grdPay.Col = 3
                        grdPay_EnterCell
                    Case 3
                        grdPay.Col = 4
                        grdPay_EnterCell
                    Case 4
                        grdPay.Col = 3
                        BFA = grdPay.Text
                        grdPay.Col = 4
                        AmountDue = grdPay.Text
                        TotalDue = BFA + AmountDue
                        grdPay.Col = 5
                        grdPay.Text = FormatCurrency(TotalDue, 2)
                        grdPay.Col = 6
                        grdPay_EnterCell
                    'Case 5
                    '    grdPay.Col = 6
                    '    grdPay_EnterCell
                    Case 6
                        grdPay.Col = 5
                        TotalDue = grdPay.Text
                        grdPay.Col = 6
                        AmountPaid = grdPay.Text
                        CFA = TotalDue - AmountPaid
                        grdPay.Col = 7
                        grdPay.Text = FormatCurrency(CFA, 2)
                        cmdSave.SetFocus
                        'grdPay_EnterCell
                    'Case 7
                    '    grdPay.Col = 8
                    '    grdPay_EnterCell
                    'Case 7
                       ' If grdPay.Col = 7 And grdPay.Text <> "" Then
                       '     grdPay.Text = FormatCurrency(grdPay.Text, 2)
                       '     grdPay.Rows = grdPay.Rows + 1
                       '     grdPay.Row = grdPay.Row + 1
                       '     grdPay.Col = 1
                       '     grdPay_EnterCell
                       ' End If
                End Select
    End Select
End Sub

Private Sub ClearGrid()
grdPay.Clear
txtEdit.Visible = False
mskEdit.Visible = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
Select Case grdPay.Col
    Case 3, 4, 5, 6, 7, 8
     KeyAscii = ValidateInput(KeyAscii, Currency_Input)
End Select

End Sub
