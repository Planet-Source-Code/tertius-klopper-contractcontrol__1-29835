VERSION 5.00
Begin VB.Form frmRepairDba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repair Database"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmRepairDba.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmRepairDba.frx":08CA
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdRepairdba 
      Caption         =   "Repair Database"
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "frmRepairDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRepairdba_Click()
Dim errLoop As Error
If MsgBox("Do you want to repair the Contract Control Database", vbYesNo) = vbYes Then
   If MsgBox("Are You sure!", vbYesNo) = vbYes Then
      DB.Close
      On Error GoTo Err_Repair
      DBEngine.RepairDatabase App.Path & "\Contract.mdb"
      OpenDB
      On Error GoTo 0
      MsgBox "End of repair procedure!"
   End If
End If

   Exit Sub
   
Err_Repair:
   For Each errLoop In DBEngine.Errors
      MsgBox "Repair unsuccessful!" & vbCr & _
         "Error number: " & errLoop.Number & _
         vbCr & errLoop.Description
   Next errLoop
End Sub

