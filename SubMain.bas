Attribute VB_Name = "SubMain"
'Programmer Tertius Klopper 06/09/2001
'Copyright © 2001 Tertius Klopper

'This program is not to be used in any commercial areas as it is only ment
'for demostration purposes. As for the source code of this program
'is ment to teach other how I did it. The code of this program may
'be used in your personal applications but any commercial application of
'the code is not permitted without written concent from programmer.
'The source code may not be distrubuted in any form, ie to diffrent
'web sites or placed on any personal web pages.
'Only site allowed to have code is Planet-Source-Code.com

'Sample Code found on PSC used in this application is
'
'Program Name - InputValidation
'Programmer   - Matt Trigwell
'Comments     - None
'
'Program Name - Credit Sales
'Programmer   - Enix Information System
'Comments     - None
'
'Program Name - BackupSheild
'Programmer   - unknown
'Comments     - None
'
Option Explicit
Global Const DEFSOURCE = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source="
Global Const DBName = "\Contract.MDB;Jet OLEDB:Database Password=matrix-se;"
Public DB As ADODB.Connection
Public Started As Boolean
Global LoginSucceeded As Boolean
Dim CheckUser As ADODB.Recordset
Dim NoUsers As Boolean
Global UserName As String
Global NoDba As Boolean
Global TotalExp As Currency
Global TotalPay As Currency
Sub Main()
If Not App.PrevInstance Then
OpenDB
If NoDba <> True Then
Set CheckUser = New ADODB.Recordset
 CheckUser.Open "SELECT * FROM userlist", DB, adOpenStatic, adLockOptimistic
    If CheckUser.RecordCount = 0 Then
        NoUsers = True
        LoginSucceeded = True
    End If
If Not NoUsers = True Then
  frmLogin.Show vbModal
End If
If Not LoginSucceeded = True Then
  End
End If
Load frmSplash
frmSplash.Show
ElseIf App.PrevInstance Then
MsgBox "There is another copy of the application already running!", vbCritical
End If
End If
End Sub
Public Sub OpenDB()
On Error GoTo err
     Set DB = New ADODB.Connection
     DB.Open DEFSOURCE & App.Path & DBName
err:
   Select Case err.Number
   Case -2147467259
   NoDba = True
   If MsgBox("Database " & App.Path & "\Contract.MDB could not be found Restore Database", vbYesNo) = vbYes Then
      frmRestoreDba.Show
   Else
   End
    End If
   End Select
 End Sub
