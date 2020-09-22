VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using DataGrid Control with DataEnvironment - Ep. 2               By: Ahmad Eissa"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "Year Born"
      DataMember      =   "Authors_Table"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Author"
      DataMember      =   "Authors_Table"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   6240
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Au_ID"
      DataMember      =   "Authors_Table"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Save Record"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<<"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Find Rec."
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Rec."
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Record"
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0442
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Authors_Table"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Au_ID"
         Caption         =   "Author ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Author"
         Caption         =   "Author Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Year Born"
         Caption         =   "Year Born"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3509.858
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Don't Forget to VOTE ME !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   17
      Top             =   7080
      Width           =   4425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Born:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub CLR()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub ENBL()
    Text1.Locked = False
    Text2.Locked = False
    Text3.Locked = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command10.Enabled = True
End Sub

Private Sub DIS()
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command10.Enabled = False
End Sub

Private Sub Command10_Click()
On Error GoTo er
    DataEnvironment1.rsAuthors_Table.UpdateBatch adAffectCurrent
    Call DIS
    DataGrid1.Refresh
    Exit Sub
er: MsgBox ("Error in entering data operation, Try Again!"), vbCritical, "ERROR"
End Sub

Private Sub Command2_Click()
    If Not (DataEnvironment1.rsAuthors_Table.EOF = True And DataEnvironment1.rsAuthors_Table.BOF = True) Then
        DataEnvironment1.rsAuthors_Table.MoveLast
    End If
    Call CLR
    Call ENBL
    DataEnvironment1.rsAuthors_Table.AddNew
End Sub

Private Sub Command3_Click()
    Call ENBL
    DataEnvironment1.rsAuthors_Table.Update
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    f = MsgBox("All Data Will Be Deleted, Are you sure?", (vbOKCancel + vbCritical), "Warning...")
    If f = vbOK Then
        DataEnvironment1.rsAuthors_Table.Delete adAffectCurrent
    End If
End Sub

Private Sub Command5_Click()
MsgBox "not in this version :)"
End Sub

Private Sub Command6_Click()
If DataEnvironment1.rsAuthors_Table.RecordCount <> 0 Then
    DataEnvironment1.rsAuthors_Table.MoveFirst
    Beep
Else: MsgBox ("No Records!"), vbInformation, Me.Caption
End If
End Sub

Private Sub Command7_Click()
If DataEnvironment1.rsAuthors_Table.RecordCount <> 0 Then
    DataEnvironment1.rsAuthors_Table.MovePrevious
    If DataEnvironment1.rsAuthors_Table.BOF Then
        MsgBox ("You are in the first record..."), vbInformation, Me.Caption
        DataEnvironment1.rsAuthors_Table.MoveFirst
    End If
Else: MsgBox ("No Records!"), vbInformation, Me.Caption
End If
End Sub

Private Sub Command8_Click()
If DataEnvironment1.rsAuthors_Table.RecordCount <> 0 Then
    DataEnvironment1.rsAuthors_Table.MoveNext
    If DataEnvironment1.rsAuthors_Table.EOF Then
        MsgBox ("You are in the last record..."), vbInformation, Me.Caption
        DataEnvironment1.rsAuthors_Table.MoveLast
    End If
Else:  MsgBox ("No Records!"), vbInformation, Me.Caption
End If
End Sub

Private Sub Command9_Click()
If DataEnvironment1.rsAuthors_Table.RecordCount <> 0 Then
    DataEnvironment1.rsAuthors_Table.MoveLast
    Beep
Else:  MsgBox ("No Records!"), vbInformation, Me.Caption
End If
End Sub

Private Sub Form_Load()
    Call DIS
End Sub
