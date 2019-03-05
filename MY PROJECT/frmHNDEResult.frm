VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHNDEResult 
   BackColor       =   &H00C0E0FF&
   Caption         =   "HND(Evening) Result Table."
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "frmHNDEResult.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   18865
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "HND I(Evening) 1st Semester"
      TabPicture(0)   =   "frmHNDEResult.frx":234CD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdBack"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Adodc1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "HND I(Evening) 2nd Semester"
      TabPicture(1)   =   "frmHNDEResult.frx":234E9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBack2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Adodc2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "HND II(Evening) 1st Semester"
      TabPicture(2)   =   "frmHNDEResult.frx":23505
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBack3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Adodc3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "HND II(Evening) 2nd Semester"
      TabPicture(3)   =   "frmHNDEResult.frx":23521
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Adodc4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "DataGrid4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdBack4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdBack4 
         Caption         =   "&Back"
         Height          =   495
         Left            =   9360
         TabIndex        =   4
         Top             =   10080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack3 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -65640
         TabIndex        =   3
         Top             =   10080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack2 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -65640
         TabIndex        =   2
         Top             =   10080
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -65640
         TabIndex        =   1
         Top             =   10080
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmHNDEResult.frx":2353D
         Height          =   8895
         Left            =   -74880
         TabIndex        =   5
         Top             =   1080
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   1000
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=StudentSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "StudentSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblHND1EFirst"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmHNDEResult.frx":23552
         Height          =   8895
         Left            =   -74880
         TabIndex        =   6
         Top             =   1080
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   1000
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=StudentSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "StudentSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblHND1eSecond"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmHNDEResult.frx":23567
         Height          =   8895
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   1000
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=StudentSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "StudentSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblHND2EFirst"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmHNDEResult.frx":2357C
         Height          =   8895
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   735
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   1000
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=StudentSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "StudentSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblHND2ESecond"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND II(EVENING) SECOND SEMESTER RESULT TABLE"
         BeginProperty Font 
            Name            =   "American Classic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Width           =   12060
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND II(EVENING) FIRST SEMESTER RESULT TABLE"
         BeginProperty Font 
            Name            =   "American Classic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -70800
         TabIndex        =   11
         Top             =   480
         Width           =   11520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND I(EVENING) SECOND SEMESTER RESULT TABLE"
         BeginProperty Font 
            Name            =   "American Classic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -70920
         TabIndex        =   10
         Top             =   480
         Width           =   11925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND I(EVENING) FIRST SEMESTER RESULT TABLE"
         BeginProperty Font 
            Name            =   "American Classic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -70680
         TabIndex        =   9
         Top             =   480
         Width           =   11385
      End
   End
End
Attribute VB_Name = "frmHNDEResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect1()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function
Private Function GetConnect2()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function
Private Function GetConnect3()
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function
Private Function GetConnect4()
Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function


Private Sub cmdBack_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 0
End Sub

Private Sub cmdBack2_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 0
End Sub

Private Sub cmdBack3_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 1
End Sub

Private Sub cmdBack4_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 1
End Sub

Private Sub Form_Load()
GetConnect1
GetConnect2
GetConnect3
GetConnect4
End Sub
