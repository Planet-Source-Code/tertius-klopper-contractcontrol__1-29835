VERSION 5.00
Begin VB.Form frmCompactDba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compact Database"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmCompactDba.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmCompactDba.frx":08CA
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdCompactdba 
      Caption         =   "Compact Database"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblFreeSpace 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblNewSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
   End
End
Attribute VB_Name = "frmCompactDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbasize As Long
Private Sub cmdCompactdba_Click()
On Error GoTo err
If MsgBox("Are you sure", vbYesNo) = vbYes Then
      DB.Close
    If Dir(App.Path & "\CompactCon.mdb") <> "" Then
      Kill App.Path & "\CompactCon.mdb"
    End If
    DBEngine.CompactDatabase App.Path & "\Contract.mdb", App.Path & "\CompactCon.mdb", , , ";pwd=matrix-se"
    Kill App.Path & "\contract.mdb"
    Name App.Path & "\CompactCon.mdb" As App.Path & "\Contract.mdb"
    PathName = App.Path & "\Contract.MDB"
    'On Error GoTo err
    dbasize = FileLen(PathName)
    lblNewSize = "Compacted Database size : " & Format((dbasize / 1024) / 1024, "standard") & "MB."
    OpenDB
End If

err:
 If err.Number = 3356 Then
   MsgBox "Error occured while trying to compact database Restart your Computer and try again", vbExclamation
   Exit Sub
End If
End Sub



Private Sub Form_Activate()
lblSize = "Current Database size: " & Format((dbasize / 1024) / 1024, "standard") & "MB."

End Sub

Private Sub Form_Load()
    Dim fs, d, s
    Dim drvpath As String
    Dim freeSpace As Long
    drvpath = App.Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvpath))
    freeSpace = d.AvailableSpace / 1024 / 1024
    s = "Drive " & Left(App.Path, 1) & " has "
    lblFreeSpace = s & FormatNumber(freeSpace, 0) & "MB free"
PathName = App.Path & "\Contract.MDB"
On Error GoTo err
dbasize = FileLen(PathName)
If freeSpace * 1024 * 1024 < dbasize Then
  lblNewSize = "Not enough space to compact database clear some space on drive " & Left(App.Path, 1)
  cmdCompactdba.Enabled = False
End If
err:
Exit Sub
End Sub

Private Sub Label1_Click()

End Sub

