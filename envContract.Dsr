VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DEnv 
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   _ExtentX        =   10874
   _ExtentY        =   11351
   FolderFlags     =   7
   TypeLibGuid     =   "{2CDD5E66-5FFC-11D5-A4ED-D0974DC13E01}"
   TypeInfoGuid    =   "{2CDD5E67-5FFC-11D5-A4ED-D0974DC13E01}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "conContract"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"envContract.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   11
   BeginProperty Recordset1 
      CommandName     =   "SubConMaxNum"
      CommDispId      =   1002
      RsDispId        =   1006
      CommandText     =   "SELECT MAX(SubConNum) AS MaxNum FROM SubConDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "MaxNum"
         Caption         =   "MaxNum"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdAllSubCon"
      CommDispId      =   1007
      RsDispId        =   1014
      CommandText     =   "SELECT * FROM SubConDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SubConNum"
         Caption         =   "SubConNum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConFname"
         Caption         =   "SubConFname"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConLname"
         Caption         =   "SubConLname"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdSubCon"
      CommDispId      =   1015
      RsDispId        =   1031
      CommandText     =   $"envContract.dsx":00B8
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SubConNum"
         Caption         =   "SubConNum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConFname"
         Caption         =   "SubConFname"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConLname"
         Caption         =   "SubConLname"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "cmdWorkDetail"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"envContract.dsx":014B
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdSubCon"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SubConNum"
         Caption         =   "SubConNum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "ContractNo"
         Caption         =   "ContractNo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "WorkDescription"
         Caption         =   "WorkDescription"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "WorkType"
         Caption         =   "WorkType"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "TransferDate"
         Caption         =   "TransferDate"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "SubConNum"
         ChildField      =   "SubConNum"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "cmdPayments"
      CommDispId      =   1032
      RsDispId        =   1037
      CommandText     =   "SELECT DISTINCT SubConDetails.* FROM SubConDetails, SubConPayments WHERE SubConDetails.SubConNum = SubConPayments.SubConNum"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SubConNum"
         Caption         =   "SubConNum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConFname"
         Caption         =   "SubConFname"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SubConLname"
         Caption         =   "SubConLname"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "cmdSubConPay"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT SubConPayments.* FROM SubConPayments, SubConDetails WHERE SubConPayments.SubConNum = SubConDetails.SubConNum"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdPayments"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "SubConNum"
         Caption         =   "SubConNum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ChequeNo"
         Caption         =   "ChequeNo"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "AmountPaid"
         Caption         =   "AmountPaid"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BFA"
         Caption         =   "BFA"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "LBF"
         Caption         =   "LBF"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "LPB"
         Caption         =   "LPB"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "AmountDue"
         Caption         =   "AmountDue"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CFA"
         Caption         =   "CFA"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "SubConNum"
         ChildField      =   "SubConNum"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "cmdConList"
      CommDispId      =   1038
      RsDispId        =   1045
      CommandText     =   "SELECT ContractNo, Amount, ClientName, Estimator, QuoteNo, Completed FROM ContractDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "ContractNo"
         Caption         =   "ContractNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ClientName"
         Caption         =   "ClientName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Estimator"
         Caption         =   "Estimator"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "QuoteNo"
         Caption         =   "QuoteNo"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Completed"
         Caption         =   "Completed"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "cmdConDetails"
      CommDispId      =   1046
      RsDispId        =   1052
      CommandText     =   "SELECT ContractDetails.* FROM ContractDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "QuoteNo"
         Caption         =   "QuoteNo"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Estimator"
         Caption         =   "Estimator"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "ContractNo"
         Caption         =   "ContractNo"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ClientName"
         Caption         =   "ClientName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "WorkDescription"
         Caption         =   "WorkDescription"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "ClientAddress"
         Caption         =   "ClientAddress"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Completed"
         Caption         =   "Completed"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "cmdConConSum"
      CommDispId      =   1053
      RsDispId        =   1059
      CommandText     =   "SELECT DISTINCT ContractNo, ClientName, Amount FROM ContractDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "ContractNo"
         Caption         =   "ContractNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ClientName"
         Caption         =   "ClientName"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "Command1"
      CommDispId      =   1060
      RsDispId        =   1063
      CommandText     =   "ContractDetails"
      ActiveConnectionName=   "conContract"
      CommandType     =   2
      dbObjectType    =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "QuoteNo"
         Caption         =   "QuoteNo"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Amount"
         Caption         =   "Amount"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Estimator"
         Caption         =   "Estimator"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "ContractNo"
         Caption         =   "ContractNo"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ClientName"
         Caption         =   "ClientName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "WorkDescription"
         Caption         =   "WorkDescription"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "ClientAddress"
         Caption         =   "ClientAddress"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Completed"
         Caption         =   "Completed"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "Command2"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   ""
      ActiveConnectionName=   "conContract"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Command1"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub rsSubConNumbers_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

