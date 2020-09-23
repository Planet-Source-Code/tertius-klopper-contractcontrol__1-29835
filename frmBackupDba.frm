VERSION 5.00
Begin VB.Form frmBackupDba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Database"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmBackupDba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup Database"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Backup Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblDbaSize 
      Alignment       =   2  'Center
      Caption         =   "Current Database Size is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmBackupDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbasize As Long
Dim PathName As String

Private Sub cmdBackup_Click()
If txtDestination <> "" Then
DoBackup PathName, txtDestination
ElseIf txtDestination = "" Then
  MsgBox "You must specify a distination for the backup", vbCritical
End If
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me

End Sub

Private Sub cmdDestination_Click()
Dim strTemp As String

strTemp = fBrowseForFolder(Me.hwnd, "Select backup path")
If strTemp <> "" Then
    txtDestination = strTemp
End If

End Sub

Private Sub Form_Activate()
lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
End Sub

Private Sub Form_Load()
'SetRegion
PathName = App.Path & "\Contract.MDB"
dbasize = FileLen(PathName)
End Sub
Private Sub SetRegion()
    On Error Resume Next
    If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(Me.Picture, RGB(255, 0, 255))
    SetWindowRgn Me.hwnd, hRgn, True
End Sub

