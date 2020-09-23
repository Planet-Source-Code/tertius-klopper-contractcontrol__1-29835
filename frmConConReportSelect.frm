VERSION 5.00
Begin VB.Form frmConConReportSelect 
   Caption         =   "Report Select"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "frmConConReportSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2018
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   578
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox cboContractNo 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Contract Number :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmConConReportSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
