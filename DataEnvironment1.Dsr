VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   _ExtentX        =   18547
   _ExtentY        =   16748
   FolderFlags     =   1
   TypeLibGuid     =   "{C3CD6EA7-751F-4506-A0B2-5F690ADD8438}"
   TypeInfoGuid    =   "{18FBD184-0983-45B9-AC17-605A15FF8E10}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"DataEnvironment1.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   1
   BeginProperty Recordset1 
      CommandName     =   "Authors_Table"
      CommDispId      =   1002
      RsDispId        =   1007
      CommandText     =   "Authors"
      ActiveConnectionName=   "Connection1"
      CommandType     =   2
      dbObjectType    =   1
      CursorType      =   2
      Locktype        =   3
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Au_ID"
         Caption         =   "Au_ID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Author"
         Caption         =   "Author"
      EndProperty
      BeginProperty Field3 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "Year Born"
         Caption         =   "Year Born"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()
    DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BIBLIO-2K.mdb;Persist Security Info=False"
End Sub
