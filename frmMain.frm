VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract Control"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8.255
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   13.07
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   1200
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrStatusbar 
      Interval        =   1000
      Left            =   720
      Top             =   3840
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4425
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5901
            MinWidth        =   4234
            Text            =   "User Name :"
            TextSave        =   "User Name :"
            Object.ToolTipText     =   "Current User Log In"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList lstMenuImages 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6562
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":687E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":755A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":815A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9316
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMainMenu 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6165
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "lstMenuImages"
      SmallIcons      =   "lstMenuImages"
      ColHdrIcons     =   "lstMenuImages"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
      Picture         =   "frmMain.frx":A4CE
   End
   Begin VB.Label lblMenuCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image imgCaption 
      Height          =   720
      Left            =   120
      Picture         =   "frmMain.frx":B998
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "&Print Setup"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSubCon 
      Caption         =   "&Subcontractors"
      Begin VB.Menu mnuSubConAdd 
         Caption         =   "&Add Subcontractor"
      End
      Begin VB.Menu mnuSubConWorkDetail 
         Caption         =   "&Work Details"
      End
      Begin VB.Menu mnuSubConPayments 
         Caption         =   "&Payments"
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "&Contracts"
      Begin VB.Menu mnuConContractDetails 
         Caption         =   "&Contract &Details"
      End
      Begin VB.Menu mnuConContractControl 
         Caption         =   "C&ontract Control"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsUserSetup 
         Caption         =   "&User Setup"
      End
      Begin VB.Menu mnuOptionsBackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnuOptionsRestoreDatabase 
         Caption         =   "&Restore Database"
      End
      Begin VB.Menu mnuOptionsRepair 
         Caption         =   "Repair &Database"
      End
      Begin VB.Menu mnuOptionsCompact 
         Caption         =   "Com&pact Database"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "C&ontents"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Contract Control"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Single, yy As Single

Dim TodayDate As Date
Private Sub Form_Load()
  StartMenu
  If UserName <> "" Then
   StatusBar.Panels(1).Text = "User Name : " & UserName
  Else
   StatusBar.Panels(1).Text = "User Name : Login Not Required"
  End If
 StatusBar.Panels(2).Text = Format(Date, "Long Date")
 StatusBar.Panels(3).Text = Time
End Sub
Private Sub StartMenu()
    Dim ItmX As ListItem
    lstMainMenu.ListItems.Clear
    Set ItmX = lstMainMenu.ListItems.Add(1, "Emp", "Subcontractors", 15)
    Set ItmX = lstMainMenu.ListItems.Add(2, "Con", "Contracts", 18)
    Set ItmX = lstMainMenu.ListItems.Add(3, "Opt", "Options", 2)
    Set ItmX = lstMainMenu.ListItems.Add(4, "Hlp", "Help", 4)
    Set ItmX = lstMainMenu.ListItems.Add(5, "Exit", "Exit", 6)
    lblMenuCaption.Caption = "Main Menu"
End Sub
Private Sub EmpMenu()
   Dim ItmX As ListItem
   lstMainMenu.ListItems.Clear
   Set ItmX = lstMainMenu.ListItems.Add(1, "AddEmp", "Add Subcontractor", 14)
   Set ItmX = lstMainMenu.ListItems.Add(2, "Work", "Work Details", 13)
   Set ItmX = lstMainMenu.ListItems.Add(3, "Pay", "Payments", 11)
   Set ItmX = lstMainMenu.ListItems.Add(4, "SubConRpt", "Reports", 1)
   'Last One In Index
   Set ItmX = lstMainMenu.ListItems.Add(5, "Main", "Main Menu", 7)
   lblMenuCaption.Caption = "SubContractors"
End Sub
Private Sub SubConRptMenu()
Dim ItmX As ListItem
lstMainMenu.ListItems.Clear
Set ItmX = lstMainMenu.ListItems.Add(1, "lstSubCon", "Subcontractor List", 1)
Set ItmX = lstMainMenu.ListItems.Add(2, "WorkDetail", "Work Details Report", 1)
Set ItmX = lstMainMenu.ListItems.Add(3, "SubConPay", "SubContractor Payments", 1)
'Last One In Index
Set ItmX = lstMainMenu.ListItems.Add(4, "Main", "Main Menu", 7)
lblMenuCaption.Caption = "Report"
End Sub
Private Sub ConMenu()
   Dim ItmX As ListItem
    lstMainMenu.ListItems.Clear
    Set ItmX = lstMainMenu.ListItems.Add(1, "AddCon", "Contract Details", 5)
    Set ItmX = lstMainMenu.ListItems.Add(2, "ConCtl", "Contract Control", 12)
    Set ItmX = lstMainMenu.ListItems.Add(3, "ConRpt", "Reports", 1)
    'Last One In Index
    Set ItmX = lstMainMenu.ListItems.Add(4, "Main", "Main Menu", 7)
    lblMenuCaption.Caption = "Contracts"
End Sub
Private Sub ConRptMenu()
Dim ItmX As ListItem
 lstMainMenu.ListItems.Clear
 Set ItmX = lstMainMenu.ListItems.Add(1, "ConDetails", "Contract Details", 1)
 'Last on in Indexs
 Set ItmX = lstMainMenu.ListItems.Add(2, "Main", "Main Menu", 7)
 lblMenuCaption.Caption = "Reports"
End Sub

Private Sub OptMenu()
   Dim ItmX As ListItem
    lstMainMenu.ListItems.Clear
    Set ItmX = lstMainMenu.ListItems.Add(1, "User", "User Setup", 10)
    Set ItmX = lstMainMenu.ListItems.Add(2, "Back", "Backup Database", 8)
    Set ItmX = lstMainMenu.ListItems.Add(3, "Restore", "Restore Database", 9)
    Set ItmX = lstMainMenu.ListItems.Add(4, "Repair", "Repair Database", 16)
    Set ItmX = lstMainMenu.ListItems.Add(5, "Compact", "Compact Database", 17)
    'Last on in Index
    Set ItmX = lstMainMenu.ListItems.Add(6, "Main", "Main Menu", 7)
    lblMenuCaption.Caption = "Options"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer
Response = MsgBox("Are You Sure", vbYesNo + vbDefaultButton2, "Exit")
If Response = vbYes Then   ' User chose Yes.
    Cancel = 0
Else   ' User chose No.
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub lstMainMenu_Click()
  LoadModules
End Sub
Private Sub lstMainMenu_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
     LoadModules
 End If
End Sub
Private Sub LoadModules()
Dim nRun
    On Error GoTo ExitThis
    Select Case lstMainMenu.HitTest(xx, yy).Key
           Case "Emp"
                  EmpMenu 'Subcontractor Menu
             Case "AddEmp"
                   frmAddSubCon.Show
             Case "Work"
                   frmSubWork.Show
             Case "Pay"
                   frmSubPayments.Show
             Case "SubConRpt"
                   SubConRptMenu 'Subcontractor Report Menu
               Case "lstSubCon"
                   rptAllSubCon.Show
               Case "WorkDetail"
                   frmSubConReportSelect.optWorkDetailsrpt.Value = True
                   frmSubConReportSelect.Show
               Case "SubConPay"
                   frmSubConReportSelect.optPaymentrpt.Value = True
                   frmSubConReportSelect.Show
           Case "Con"
                  ConMenu 'Contract Menu
             Case "AddCon"
                  frmAddContract.Show
             Case "ConCtl"
                  frmContractControl.Show
             Case "ConRpt" 'Contract Report Menu
                  ConRptMenu
               Case "ConDetails"
                  frmConReportSelect.optConDetails.Value = True
                  frmConReportSelect.Show
           Case "Opt"
                  OptMenu
             Case "User"
                 frmUserSetup.Show
             Case "Back"
                 frmBackupDba.Show
             Case "Restore"
                 frmRestoreDba.Show
             Case "Repair"
                 frmRepairDba.Show
             Case "Compact"
                 frmCompactDba.Show
           Case "Hlp"
           comDialog.HelpFile = App.Path & "\Help\ContractControl.hlp"
           comDialog.HelpCommand = &HB 'cdlHelpContents
           comDialog.ShowHelp   ' Display Visual Basic Help contents topic.
      
           Case "Main"
               StartMenu
           Case "Exit"
               Unload frmMain
    End Select
    
ExitThis:
End Sub

Private Sub lstMainMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 xx = X
 yy = y
End Sub

Private Sub mnuConContractControl_Click()
frmContractControl.Show
End Sub

Private Sub mnuConContractDetails_Click()
frmAddContract.Show
End Sub

Private Sub mnuFileExit_Click()
Unload frmMain
End Sub

Private Sub mnuFilePrinterSetup_Click()
comDialog.ShowPrinter
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuHelpContents_Click()
comDialog.HelpFile = App.Path & "\Help\ContractControl.hlp"
comDialog.HelpCommand = &HB 'cdlHelpContents
comDialog.ShowHelp   ' Display Visual Basic Help contents topic.
End Sub
Private Sub mnuOptionsBackup_Click()
frmBackupDba.Show
End Sub

Private Sub mnuOptionsCompact_Click()
frmCompactDba.Show
End Sub

Private Sub mnuOptionsRepair_Click()
frmRepairDba.Show
End Sub

Private Sub mnuOptionsRestoreDatabase_Click()
frmRestoreDba.Show
End Sub

Private Sub mnuOptionsUserSetup_Click()
frmUserSetup.Show
End Sub

Private Sub mnuSubConAdd_Click()
frmAddSubCon.Show
End Sub

Private Sub mnuSubConPayments_Click()
frmSubPayments.Show
End Sub

Private Sub mnuSubConWorkDetail_Click()
frmSubWork.Show
End Sub

Private Sub tmrStatusbar_Timer()
StatusBar.Panels(2).Text = Format(Date, "Long Date")
StatusBar.Panels(3).Text = Time
End Sub
