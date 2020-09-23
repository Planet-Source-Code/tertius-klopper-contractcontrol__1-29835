VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontracto Work Details"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmSubWork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Click to Export Work Details"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   135
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
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
   Begin VB.ListBox lstWorkType 
      Height          =   255
      IntegralHeight  =   0   'False
      ItemData        =   "frmSubWork.frx":08CA
      Left            =   2520
      List            =   "frmSubWork.frx":08D4
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWork 
      Height          =   3495
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      _Version        =   393216
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frmViewList 
      Caption         =   "View Lists"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin VB.OptionButton optViewJobCon 
         Caption         =   "View Jobs On Specific Contract"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Double Click to Enter Contract Number"
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton optViewAll 
         Caption         =   "View All Jobs"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         ToolTipText     =   "Click to View All Details"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optInsert 
         Caption         =   "Insert New Work Details"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Click to Insert New Work Details"
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComctlLib.ListView lstSubCon 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7858
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
Attribute VB_Name = "frmSubWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LSubRs As ADODB.Recordset
Dim SubDetails As ADODB.Recordset
Dim CheckConNum As ADODB.Recordset
Dim SubConRs As ADODB.Recordset
Dim sqlString As String
Dim InputNum As String
Dim ContractNum As String

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim SubConDel As ADODB.Recordset
Dim ContNum As String
Dim ConDate As String
If SubConRs.RecordCount <> 0 Then
         SubConRs.MoveFirst
    If Not SubConRs.EOF Or SubConRs.BOF Then
        If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
         Set SubConDel = New ADODB.Recordset
         grdWork.Row = grdWork.Row
         grdWork.Col = 1: ContNum = grdWork.Text
         grdWork.Col = 2: ConDate = grdWork.Text
         sqlString = "Select * from SubConWorkDetails where SubConNum=" & _
         Trim(lstSubCon.SelectedItem.SubItems(1)) & _
         " and ContractNo='" & Trim(ContNum) & "' and Date=#" & Format(ConDate, "m/d/yyyy") & "#"
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

Private Sub cmdExport_Click()
frmExport.Show
End Sub

Private Sub cmdPrint_Click()
frmSubConReportSelect.optWorkDetailsrpt.Value = True
frmSubConReportSelect.Show

End Sub

Private Sub cmdSave_Click()
Dim IDate
Dim r As Integer
     Set SubDetails = New ADODB.Recordset
     SubDetails.Open "Select * from subconworkdetails", DB, adOpenStatic, adLockOptimistic
     For r = 1 To grdWork.Rows - 2
        grdWork.Row = r
        With SubDetails
            .AddNew
            !SubConNum = lstSubCon.SelectedItem.SubItems(1)
            grdWork.Col = 1: !ContractNo = grdWork.Text
            grdWork.Col = 2: IDate = grdWork.Text
            grdWork.Col = 2: !Date = Format(IDate, "dd/mm/yyyy")
            grdWork.Col = 3: !WorkDescription = grdWork.Text
            grdWork.Col = 4: !Worktype = grdWork.Text
            grdWork.Col = 5: !Amount = grdWork.Text
            .Update
        End With
     Next
     ClearGrid
     CreateGrid
End Sub

Private Sub cmdUpdate_Click()
Dim IDate
Dim r As Integer
Dim SubNum As String
     For r = 1 To grdWork.Rows - 1
     grdWork.Row = r
     grdWork.Col = 1: SubNum = grdWork.Text
     grdWork.Col = 2: IDate = grdWork.Text
     Set SubDetails = New ADODB.Recordset
     sqlString = "Select * from SubConWorkDetails where SubConNum=" & _
     Trim(lstSubCon.SelectedItem.SubItems(1)) & _
     " and ContractNo='" & Trim(SubNum) & "' and Date=#" & Format(IDate, "m/d/yyyy") & "#"
     SubDetails.Open sqlString, DB, adOpenStatic, adLockOptimistic
     If SubDetails.RecordCount <> 0 Then
        With SubDetails
            !SubConNum = lstSubCon.SelectedItem.SubItems(1)
            !ContractNo = SubNum
            !Date = CDate(IDate)
            grdWork.Col = 3: !WorkDescription = grdWork.Text
            grdWork.Col = 4: !Worktype = grdWork.Text
            grdWork.Col = 5: !Amount = grdWork.Text
            .Update
        End With
     End If
     Next
     ClearGrid
     FromListUpdateRecord
End Sub

Private Sub Form_Activate()
optInsert.SetFocus
lstSubCon.SetFocus
End Sub

Private Sub Form_Load()
grdWork.Cols = 6
grdWork.Rows = grdWork.Rows
grdWork.Row = 1
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

Private Sub FromListUpdateRecord()
Dim r As Integer

If optViewAll.Value = True Then
 Set SubConRs = New ADODB.Recordset
 If Not lstSubCon.SelectedItem.Text = Empty Then
   sqlString = Trim(lstSubCon.SelectedItem.SubItems(1))
   SubConRs.Open "Select * from SubconWorkDetails where SubconNum=" & sqlString, DB, adOpenStatic, adLockOptimistic
     grdWork.Clear
     grdWork.Cols = 6
     grdWork.Rows = 2
     grdWork.Row = 0
     CreateGrid
   If SubConRs.RecordCount <> 0 Then
    prgLoad.Max = SubConRs.RecordCount
   Else
    prgLoad.Max = 1
   End If
   grdWork.Row = grdWork.Row
   grdWork.Col = 1
    If SubConRs.RecordCount <> 0 Then
      For r = 1 To SubConRs.RecordCount
       With SubConRs
         grdWork.Rows = grdWork.Rows + 1
         prgLoad.Value = r
         grdWork.Col = 1: grdWork.Text = !ContractNo
         grdWork.Col = 2: grdWork.Text = Format(!Date, "dd/mm/yyyy")
         grdWork.Col = 3: grdWork.Text = !WorkDescription
         grdWork.Col = 4: grdWork.Text = !Worktype
         grdWork.Col = 5: grdWork.Text = FormatCurrency(!Amount, 2)
         grdWork.Row = grdWork.Row + 1
         grdWork.Col = 1
         .MoveNext
       End With
      Next
    End If
    If grdWork.Rows > 2 Then
    grdWork.Rows = grdWork.Rows - 1
    End If
 End If

ElseIf optViewJobCon = True Then
 Set SubConRs = New ADODB.Recordset
 If Not lstSubCon.SelectedItem.Text = Empty Then
   sqlString = Trim(lstSubCon.SelectedItem.SubItems(1)) & " and contractno='" & Trim(InputNum) & "'"
   SubConRs.Open "Select * from SubconWorkDetails where SubconNum=" & sqlString, DB, adOpenStatic, adLockOptimistic
     grdWork.Clear
     grdWork.Cols = 6
     grdWork.Rows = 2
     grdWork.Row = 0
     CreateGrid
   If SubConRs.RecordCount <> 0 Then
    prgLoad.Max = SubConRs.RecordCount
   Else
    prgLoad.Max = 1
   End If
   grdWork.Row = grdWork.Row
   grdWork.Col = 1
    If SubConRs.RecordCount <> 0 Then
      For r = 1 To SubConRs.RecordCount
       With SubConRs
         grdWork.Rows = grdWork.Rows + 1
         prgLoad.Value = r
         grdWork.Col = 1: grdWork.Text = !ContractNo
         grdWork.Col = 2: grdWork.Text = Format(!Date, "dd/mm/yyyy")
         grdWork.Col = 3: grdWork.Text = !WorkDescription
         grdWork.Col = 4: grdWork.Text = !Worktype
         grdWork.Col = 5: grdWork.Text = FormatCurrency(!Amount, 2)
         grdWork.Row = grdWork.Row + 1
         grdWork.Col = 1
         .MoveNext
       End With
      Next
    End If
    If grdWork.Rows > 2 Then
    grdWork.Rows = grdWork.Rows - 1
    End If

 End If



End If
End Sub

Private Sub grdWork_EnterCell()
    '// when click on cell
 'If grdWork.Rows >= 2  Or optInsert = True Then
    Select Case grdWork.Col
        Case 1
            If optInsert.Value = True Then
            With txtEdit
                .Move grdWork.CellLeft + grdWork.Left, _
                grdWork.CellTop + grdWork.Top, grdWork.CellWidth - 25, _
                grdWork.CellHeight - 25
                .Text = grdWork.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
            ElseIf optInsert.Value = False Then
            grdWork.Col = 3
            grdWork_EnterCell
            End If
        Case 2
            If optInsert.Value = True Then
            With mskEdit
             .Move grdWork.CellLeft + grdWork.Left, _
             grdWork.CellTop + grdWork.Top, grdWork.CellWidth - 25, _
             grdWork.CellHeight - 25
               .SelText = grdWork.Text
              If Len(.Text) > 0 Then
                 .SelStart = 0
                 .SelLength = Len(.Text)
                  
              End If
              .Visible = True
              .ZOrder 0
              .SetFocus
           End With
           ElseIf optInsert.Value = False Then
           grdWork.Col = 3
           grdWork_EnterCell
           End If
        Case 3, 5
            With txtEdit
              .Move grdWork.CellLeft + grdWork.Left, _
              grdWork.CellTop + grdWork.Top, grdWork.CellWidth - 25, _
              grdWork.CellHeight - 25
              .Text = grdWork.Text
              If Len(.Text) > 0 Then
                  .SelStart = 0
                  .SelLength = Len(.Text)
              End If
               .Visible = True
               .ZOrder 0
               .SetFocus
            End With
        Case 4
          With lstWorkType
             .Move grdWork.CellLeft + grdWork.Left, _
             grdWork.CellTop + grdWork.Top, grdWork.CellWidth - 25, _
             grdWork.CellHeight - 25
             .Text = grdWork.Text
             .Visible = True
             .ZOrder 0
             .SetFocus
          End With
             
    End Select
  'End If

End Sub

Private Sub lstSubCon_Click()
prgLoad.Value = 0
FromListUpdateRecord
txtEdit.Visible = False
mskEdit.Visible = False
lstWorkType.Visible = False
End Sub
Private Sub lstSubCon_KeyUp(KeyCode As Integer, Shift As Integer)
prgLoad.Value = 0
FromListUpdateRecord
txtEdit.Visible = False
mskEdit.Visible = False
lstWorkType.Visible = False
End Sub
Private Sub lstWorkType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With lstWorkType
                    .Text = Empty
                    .Visible = False
                End With
                grdWork.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With lstWorkType
                    If .Text = Empty Then
                    .Text = "Normal"
                    End If
                    If Not .Text = Empty Then
                        grdWork.Text = .Text
                        grdWork.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                Select Case grdWork.Col
                    Case 1
                        grdWork.Col = 2
                        grdWork_EnterCell
                    Case 2
                        grdWork.Col = 3
                        grdWork_EnterCell
                    Case 3
                        grdWork.Col = 4
                        grdWork_EnterCell
                    Case 4
                        grdWork.Col = 5
                        grdWork_EnterCell
                    Case 5
                        If grdWork.Col = 5 And grdWork.Text <> "" Then
                            grdWork.Rows = grdWork.Rows + 1
                            grdWork.Row = grdWork.Row + 1
                            grdWork.Col = 1
                            grdWork_EnterCell
                         End If
                End Select
    End Select
End Sub

Private Sub mskEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CheckRs As ADODB.Recordset
Dim ConNum As String
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With txtEdit
                    .Text = Empty
                    .Visible = False
                End With
                grdWork.SetFocus
            Case vbKeyReturn
                '// when enter is pressed, move to next col
                With mskEdit
                    If Not .Text = Empty Then
                        If Not IsDate(.Text) Then
                          MsgBox "Not a valid date"
                          .SetFocus
                          Exit Sub
                        Else
                        Set CheckRs = New ADODB.Recordset
                        sqlString = "Select * from SubConWorkDetails where SubConNum=" & _
                        Trim(lstSubCon.SelectedItem.SubItems(1)) & _
                        " and ContractNo='" & Trim(ContractNum) & "' and Date=#" & Format(.Text, "m/d/yyyy") & "#"
                        CheckRs.Open sqlString, DB, adOpenStatic, adLockOptimistic
                         If CheckRs.RecordCount <> 0 Then
                          MsgBox "Cannot have two entries with the same ContractNo and Date for the selected person"
                          .SetFocus
                          Exit Sub
                         End If
                        grdWork.Text = .Text
                        End If
                    End If
                    .Visible = False
                    .SelText = Empty
                End With
                Select Case grdWork.Col
                    Case 1
                        grdWork.Col = 2
                        grdWork_EnterCell
                    Case 2
                        grdWork.Col = 3
                        grdWork_EnterCell
                    Case 3
                        grdWork.Col = 4
                        grdWork_EnterCell
                    Case 4
                        grdWork.Col = 5
                        grdWork_EnterCell
                    Case 5
                        If grdWork.Col = 5 And grdWork.Text <> "" Then
                            grdWork.Rows = grdWork.Rows + 1
                            grdWork.Row = grdWork.Row + 1
                            grdWork.Col = 1
                            grdWork_EnterCell
                         End If
                End Select
    End Select
End Sub

Private Sub optInsert_Click()
ClearGrid
CreateGrid
cmdSave.Enabled = True
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmdPrint.Enabled = False
End Sub

Private Sub optViewAll_Click()
ClearGrid
CreateGrid
cmdSave.Enabled = False
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
cmdPrint.Enabled = True
End Sub

Private Sub optViewJobCon_Click()
ClearGrid
CreateGrid
cmdSave.Enabled = False
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
'InputNum = InputBox("Enter the Contract Number", "View job's on specific contract")
cmdPrint.Enabled = True
End Sub

Private Sub optViewJobCon_DblClick()
ClearGrid
CreateGrid
cmdSave.Enabled = False
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
InputNum = InputBox("Enter the Contract Number", "View job's on specific contract")
cmdPrint.Enabled = True
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ConNum As ADODB.Recordset
    Select Case KeyCode
            Case vbKeyEscape
                '// when esc is pressed cancel and get out
                With txtEdit
                    .Text = Empty
                    .Visible = False
                End With
                grdWork.SetFocus
          Case vbKeyReturn
                '// when enter is pressed, move to next col
                With txtEdit
                    If .Text = Empty And grdWork.Col = 1 Then
                     MsgBox "You must supply a Contract Number"
                     .SetFocus
                     Exit Sub
                    End If
                    If Not .Text = Empty Then
                       Select Case grdWork.Col
                         Case 1
                           Set ConNum = New ADODB.Recordset
                           sqlString = "Select contractno from contractdetails  where contractno ='" & Trim(.Text) & "'"
                           ConNum.Open sqlString, DB, adOpenStatic, adLockOptimistic
                           If ConNum.RecordCount = 0 Then
                            MsgBox "This contract number does not exist"
                            .SetFocus
                            Exit Sub
                           End If
                           ContractNum = .Text
                       End Select
                      grdWork.Text = .Text
                    End If
                    .Visible = False
                    .Text = Empty
                End With
                Select Case grdWork.Col
                    Case 1
                        grdWork.Col = 2
                        grdWork_EnterCell
                    Case 2
                        grdWork.Col = 3
                        grdWork_EnterCell
                    Case 3
                        grdWork.Col = 4
                        grdWork_EnterCell
                    Case 4
                        grdWork.Col = 5
                        grdWork_EnterCell
                    Case 5
                        If grdWork.Col = 5 And grdWork.Text <> "" And optInsert.Value = True Then
                            grdWork.Text = FormatCurrency(grdWork.Text, 2)
                            grdWork.Rows = grdWork.Rows + 1
                            grdWork.Row = grdWork.Row + 1
                            grdWork.Col = 1
                            grdWork_EnterCell
                         ElseIf grdWork.Row <> grdWork.Rows - 1 Then
                            grdWork.Text = FormatCurrency(grdWork.Text, 2)
                            grdWork.Row = grdWork.Row + 1
                            grdWork.Col = 1
                            grdWork_EnterCell
                         Else
                            grdWork.Text = FormatCurrency(grdWork.Text, 2)
                         End If
                End Select
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If grdWork.Col = 5 Then
     KeyAscii = ValidateInput(KeyAscii, Currency_Input)
End If
End Sub

Private Sub CreateGrid()
grdWork.Col = 1: grdWork.Text = "Contract Num"
grdWork.Col = 2: grdWork.Text = "Date"
grdWork.Col = 3: grdWork.Text = "Work Description"
grdWork.Col = 4: grdWork.Text = "Work Type"
grdWork.Col = 5: grdWork.Text = "Amount"
grdWork.ColWidth(0) = 300
grdWork.ColWidth(1) = Len("Contract Num") * 100
grdWork.ColWidth(2) = Len("12/12/2001") * 100
grdWork.ColWidth(3) = Len("Work Description") * 100 + 1000
grdWork.ColWidth(4) = Len("Work Type") * 100 + 250
grdWork.ColWidth(5) = Len("Amount") * 100 + 450
End Sub

Private Sub ClearGrid()
txtEdit.Visible = False
mskEdit.Visible = False
lstWorkType.Visible = False
grdWork.Clear
grdWork.Cols = 6
grdWork.Rows = 2
grdWork.Row = 0
End Sub
