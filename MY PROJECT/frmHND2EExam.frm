VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHND2EExam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HND II(Evening) Student Examination Details"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   Icon            =   "frmHND2EExam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHND2EExam 
      BackColor       =   &H00C0FFC0&
      Height          =   9375
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   7935
      Begin TabDlg.SSTab SSTab1 
         Height          =   7575
         Left            =   0
         TabIndex        =   53
         Top             =   1800
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   13361
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "First Semester"
         TabPicture(0)   =   "frmHND2EExam.frx":234CD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label6"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label7"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label8"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label9"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label10"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label11"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label12"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label13"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label14"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label32"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label31"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label30"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label29"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label28"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label27"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label24"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label23"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label66"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label67"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "adoHND2EExam1s"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "cmdAdd"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "cmdDelete"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "cmdSearch"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "cmdBack"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).ControlCount=   27
         TabCaption(1)   =   "Second Semester"
         TabPicture(1)   =   "frmHND2EExam.frx":234E9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label15"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label16"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label17"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label18"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label19"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label20"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label21"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label22"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label25"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label26"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label34"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label58"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label59"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Label60"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Label61"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Label62"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Label63"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Label64"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Label65"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "lblName2"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "adoHND2EExam2s"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "cmdBack2"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "cmdSearch2"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "cmdDelete2"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "cmdAdd2"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).ControlCount=   27
         Begin VB.CommandButton cmdAdd2 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   -73800
            TabIndex        =   93
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete2 
            Caption         =   "&Delete Rec."
            Height          =   495
            Left            =   -72480
            TabIndex        =   92
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdSearch2 
            Caption         =   "&Search Rec."
            Height          =   495
            Left            =   -71160
            TabIndex        =   91
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack2 
            Caption         =   "&Back"
            Height          =   495
            Left            =   -69840
            TabIndex        =   90
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            Caption         =   "&Back"
            Height          =   495
            Left            =   5280
            TabIndex        =   77
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search Rec."
            Height          =   495
            Left            =   3960
            TabIndex        =   76
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete Rec."
            Height          =   495
            Left            =   2640
            TabIndex        =   75
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   1320
            TabIndex        =   74
            Top             =   6960
            Width           =   1215
         End
         Begin MSAdodcLib.Adodc adoHND2EExam1s 
            Height          =   615
            Left            =   960
            Top             =   6000
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1085
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   2
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc adoHND2EExam2s 
            Height          =   615
            Left            =   -74160
            Top             =   6120
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1085
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   2
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
            RecordSource    =   "tblHND2Second"
            Caption         =   ""
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
         Begin VB.Label Label67 
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Names"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   105
            Top             =   2520
            Width           =   5055
         End
         Begin VB.Label Label66 
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "RegNo"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   104
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label lblName2 
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            DataField       =   "Names"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73080
            TabIndex        =   103
            Top             =   2520
            Width           =   5055
         End
         Begin VB.Label Label65 
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            DataField       =   "RegNo"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73080
            TabIndex        =   102
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label64 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   -70320
            TabIndex        =   101
            Top             =   4920
            Width           =   225
         End
         Begin VB.Label Label63 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   -70320
            TabIndex        =   100
            Top             =   4320
            Width           =   225
         End
         Begin VB.Label Label62 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM422"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   99
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label61 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "EED413"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   98
            Top             =   5520
            Width           =   615
         End
         Begin VB.Label Label60 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM429"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   97
            Top             =   5160
            Width           =   615
         End
         Begin VB.Label Label59 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM426"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   96
            Top             =   4800
            Width           =   615
         End
         Begin VB.Label Label58 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM424"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   95
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label34 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM423"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   94
            Top             =   4080
            Width           =   615
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74400
            TabIndex        =   89
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grading"
            BeginProperty Font 
               Name            =   "Old English Text MT"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72240
            TabIndex        =   88
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EED413"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   87
            Top             =   5520
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM429"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   86
            Top             =   5160
            Width           =   1185
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM426"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   85
            Top             =   4800
            Width           =   1185
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM424"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   84
            Top             =   4440
            Width           =   1185
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM423"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   83
            Top             =   4080
            Width           =   1185
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM422"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74160
            TabIndex        =   82
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -71160
            TabIndex        =   81
            Top             =   4320
            Width           =   765
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GPA:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -71160
            TabIndex        =   80
            Top             =   4920
            Width           =   720
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74175
            TabIndex        =   79
            Top             =   2520
            Width           =   870
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "HND II(EVENING) EXAMINATION DETAILS SECOND SEMESTER"
            BeginProperty Font 
               Name            =   "Agency FB"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   -73680
            TabIndex        =   78
            Top             =   600
            Width           =   5205
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   4800
            TabIndex        =   73
            Top             =   4920
            Width           =   225
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   4800
            TabIndex        =   72
            Top             =   4320
            Width           =   225
         End
         Begin VB.Label Label27 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM412"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   71
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label28 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "STA411"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   70
            Top             =   5520
            Width           =   615
         End
         Begin VB.Label Label29 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM416"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   69
            Top             =   5160
            Width           =   615
         End
         Begin VB.Label Label30 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM415"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   68
            Top             =   4800
            Width           =   615
         End
         Begin VB.Label Label31 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM414"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   67
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label32 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM413"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   66
            Top             =   4080
            Width           =   615
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   600
            TabIndex        =   65
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grading"
            BeginProperty Font 
               Name            =   "Old English Text MT"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2760
            TabIndex        =   64
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA411"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   63
            Top             =   5520
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM416"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   62
            Top             =   5160
            Width           =   1185
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM415"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   61
            Top             =   4800
            Width           =   1185
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM414"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   60
            Top             =   4440
            Width           =   1185
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM413"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   59
            Top             =   4080
            Width           =   1185
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM412"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   58
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3960
            TabIndex        =   57
            Top             =   4320
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GPA:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3960
            TabIndex        =   56
            Top             =   4920
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "HND II(EVENING) EXAMINATION DETAILS FIRST SEMESTER"
            BeginProperty Font 
               Name            =   "Agency FB"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   1320
            TabIndex        =   55
            Top             =   600
            Width           =   5205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   54
            Top             =   2520
            Width           =   870
         End
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   0
         Picture         =   "frmHND2EExam.frx":23505
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.Frame fraHND2EEEdit 
      BackColor       =   &H0080FFFF&
      Height          =   9375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   7935
      Begin TabDlg.SSTab SSTab2 
         Height          =   7575
         Left            =   0
         TabIndex        =   23
         Top             =   1800
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   13361
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "First Semester"
         TabPicture(0)   =   "frmHND2EExam.frx":472D3
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label33"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label35"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label36"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label37"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label38"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label39"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label40"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label41"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label42"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label43"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblGpa"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblTotal"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label44"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label45"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtRegNo"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmdCompute"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "cmdAddR"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Combo2"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Combo3"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Combo4"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Combo5"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Combo6"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Combo1"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "cmdBk2Exam"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtName"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).ControlCount=   25
         TabCaption(1)   =   "Second Semester"
         TabPicture(1)   =   "frmHND2EExam.frx":472EF
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label46"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label47"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label48"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label49"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label50"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label51"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label52"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label53"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label54"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label55"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "lblGpa2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lblTotal2"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label56"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label57"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtRegNo2"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "cmdCompute2"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "cmdAddR2"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Combo7"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Combo8"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Combo9"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Combo10"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "Combo11"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Combo12"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "cmdBk2Exam2"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "txtName2"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).ControlCount=   25
         Begin VB.TextBox txtName2 
            DataField       =   "Names"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1800
            TabIndex        =   1
            Top             =   2520
            Width           =   5175
         End
         Begin VB.CommandButton cmdBk2Exam2 
            Caption         =   "&Back"
            Height          =   495
            Left            =   4560
            TabIndex        =   10
            Top             =   6600
            Width           =   1215
         End
         Begin VB.ComboBox Combo12 
            DataField       =   "COM422"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":4730B
            Left            =   2160
            List            =   "frmHND2EExam.frx":4732A
            TabIndex        =   7
            Text            =   "..."
            Top             =   5520
            Width           =   1215
         End
         Begin VB.ComboBox Combo11 
            DataField       =   "COM429"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":4734C
            Left            =   2160
            List            =   "frmHND2EExam.frx":4736B
            TabIndex        =   6
            Text            =   "..."
            Top             =   5160
            Width           =   1215
         End
         Begin VB.ComboBox Combo10 
            DataField       =   "COM429"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":4738D
            Left            =   2160
            List            =   "frmHND2EExam.frx":473AC
            TabIndex        =   5
            Text            =   "..."
            Top             =   4800
            Width           =   1215
         End
         Begin VB.ComboBox Combo9 
            DataField       =   "COM426"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":473CE
            Left            =   2160
            List            =   "frmHND2EExam.frx":473ED
            TabIndex        =   4
            Text            =   "..."
            Top             =   4440
            Width           =   1215
         End
         Begin VB.ComboBox Combo8 
            DataField       =   "COM424"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":4740F
            Left            =   2160
            List            =   "frmHND2EExam.frx":4742E
            TabIndex        =   3
            Text            =   "..."
            Top             =   4080
            Width           =   1215
         End
         Begin VB.ComboBox Combo7 
            DataField       =   "COM423"
            DataSource      =   "adoHND2EExam2s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":47450
            Left            =   2160
            List            =   "frmHND2EExam.frx":4746F
            TabIndex        =   2
            Text            =   "..."
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddR2 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   3240
            TabIndex        =   9
            Top             =   6600
            Width           =   1215
         End
         Begin VB.CommandButton cmdCompute2 
            Caption         =   "&Commpute GPA"
            Height          =   495
            Left            =   1560
            TabIndex        =   8
            Top             =   6600
            Width           =   1575
         End
         Begin VB.TextBox txtRegNo2 
            DataField       =   "RegNo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "99/9999/xx"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   0
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            DataField       =   "Names"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -73185
            TabIndex        =   12
            Top             =   2520
            Width           =   5175
         End
         Begin VB.CommandButton cmdBk2Exam 
            Caption         =   "&Back"
            Height          =   495
            Left            =   -70440
            TabIndex        =   21
            Top             =   6600
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "COM412"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":47491
            Left            =   -72840
            List            =   "frmHND2EExam.frx":474B0
            TabIndex        =   13
            Text            =   "..."
            Top             =   3720
            Width           =   1215
         End
         Begin VB.ComboBox Combo6 
            DataField       =   "STA411"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":474D2
            Left            =   -72840
            List            =   "frmHND2EExam.frx":474F1
            TabIndex        =   18
            Text            =   "..."
            Top             =   5520
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            DataField       =   "COM416"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":47513
            Left            =   -72840
            List            =   "frmHND2EExam.frx":47532
            TabIndex        =   17
            Text            =   "..."
            Top             =   5160
            Width           =   1215
         End
         Begin VB.ComboBox Combo4 
            DataField       =   "COM415"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":47554
            Left            =   -72840
            List            =   "frmHND2EExam.frx":47573
            TabIndex        =   16
            Text            =   "..."
            Top             =   4800
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "COM414"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":47595
            Left            =   -72840
            List            =   "frmHND2EExam.frx":475B4
            TabIndex        =   15
            Text            =   "..."
            Top             =   4440
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "COM413"
            DataSource      =   "adoHND2EExam1s"
            Height          =   315
            ItemData        =   "frmHND2EExam.frx":475D6
            Left            =   -72840
            List            =   "frmHND2EExam.frx":475F5
            TabIndex        =   14
            Text            =   "..."
            Top             =   4080
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddR 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   -71760
            TabIndex        =   20
            Top             =   6600
            Width           =   1215
         End
         Begin VB.CommandButton cmdCompute 
            Caption         =   "&Commpute GPA"
            Height          =   495
            Left            =   -73440
            TabIndex        =   19
            Top             =   6600
            Width           =   1575
         End
         Begin VB.TextBox txtRegNo 
            DataField       =   "RegNo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "99/9999/xx"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoHND2EExam1s"
            Height          =   495
            Left            =   -73200
            TabIndex        =   11
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "HND II(EVENING) EXAMINATION DETAILS SECOND SEMESTER"
            BeginProperty Font 
               Name            =   "Agency FB"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   1200
            TabIndex        =   51
            Top             =   600
            Width           =   5205
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   705
            TabIndex        =   50
            Top             =   2520
            Width           =   870
         End
         Begin VB.Label lblTotal2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4560
            TabIndex        =   49
            Top             =   4560
            Width           =   1935
         End
         Begin VB.Label lblGpa2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoHND2EExam2s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4560
            TabIndex        =   48
            Top             =   5160
            Width           =   1935
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GPA:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3720
            TabIndex        =   47
            Top             =   5160
            Width           =   720
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3720
            TabIndex        =   46
            Top             =   4560
            Width           =   765
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM422"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   45
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM423"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   44
            Top             =   4080
            Width           =   1185
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM424"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   43
            Top             =   4440
            Width           =   1185
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM426"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   42
            Top             =   4800
            Width           =   1185
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM429"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   41
            Top             =   5160
            Width           =   1185
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EED413"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   720
            TabIndex        =   40
            Top             =   5520
            Width           =   1050
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grading"
            BeginProperty Font 
               Name            =   "Old English Text MT"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2640
            TabIndex        =   39
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   38
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   37
            Top             =   2520
            Width           =   870
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "HND II(EVENING) EXAMINATION DETAILS FIRST SEMESTER"
            BeginProperty Font 
               Name            =   "Agency FB"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   -73800
            TabIndex        =   36
            Top             =   600
            Width           =   5205
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   -70320
            TabIndex        =   35
            Top             =   4560
            Width           =   1935
         End
         Begin VB.Label lblGpa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoHND2EExam1s"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   -70320
            TabIndex        =   34
            Top             =   5160
            Width           =   1935
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GPA:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -71160
            TabIndex        =   33
            Top             =   5160
            Width           =   720
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -71160
            TabIndex        =   32
            Top             =   4560
            Width           =   765
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM412"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   31
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM413"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   30
            Top             =   4080
            Width           =   1185
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM414"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   29
            Top             =   4440
            Width           =   1185
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM415"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   28
            Top             =   4800
            Width           =   1185
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM416"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   27
            Top             =   5160
            Width           =   1185
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA411"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74280
            TabIndex        =   26
            Top             =   5520
            Width           =   1050
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grading"
            BeginProperty Font 
               Name            =   "Old English Text MT"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   25
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No:"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74520
            TabIndex        =   24
            Top             =   1920
            Width           =   1080
         End
      End
      Begin VB.Image Image2 
         Height          =   1815
         Left            =   0
         Picture         =   "frmHND2EExam.frx":47617
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmHND2EExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const totalchr As Single = 30
Const totalchr2 As Single = 25

Private Function GetConnect1()
adoHND2EExam1s.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function
Private Function GetConnect2()
adoHND2EExam2s.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function

Private Sub cmdAdd_Click()
fraHND2EExam.Visible = False
fraHND2EEEdit.Visible = True
SSTab2.Tab = 0
End Sub

Private Sub cmdAdd2_Click()
fraHND2EExam.Visible = False
fraHND2EEEdit.Visible = True
SSTab2.Tab = 1
End Sub


Private Sub cmdBack_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 1
End Sub

Private Sub cmdBack2_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 1
End Sub

Private Sub cmdBk2Exam_Click()
fraHND2EEEdit.Visible = False
fraHND2EExam.Visible = True
SSTab1.Tab = 0
End Sub

Private Sub cmdBk2Exam2_Click()
fraHND2EEEdit.Visible = False
fraHND2EExam.Visible = True
SSTab1.Tab = 1
End Sub

Private Sub cmdDelete_Click()
On Error GoTo joe
GetConnect1
don = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, "WARNING")
If don = vbYes Then
With adoHND2EExam1s.Recordset
.Delete
.MoveNext
If .EOF Then
.MoveLast
End If
End With
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdDelete2_Click()
On Error GoTo joe
GetConnect2
don2 = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, "WARNING")
If don2 = vbYes Then
With adoHND2EExam2s.Recordset
.Delete
.MoveNext
If .EOF Then
.MoveLast
End If
End With
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdSearch_Click()
On Error GoTo joe
GetConnect1
Dim con As String
con = InputBox("Enter Student Reg. Number", "Search By Reg. No.")
BookMark1 = adoHND2EExam1s.Recordset.Bookmark
adoHND2EExam1s.Recordset.MoveFirst
adoHND2EExam1s.Recordset.Find "regno = '" & con & "'", 0, adSearchForward
If adoHND2EExam1s.Recordset.EOF = True Then
adoHND2EExam1s.Recordset.Bookmark = BookMark1
MsgBox ("No Record Found")
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdAddR_Click()
On Error GoTo joe
GetConnect1
adoHND2EExam1s.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdAddR2_Click()
On Error GoTo joe
GetConnect2
adoHND2EExam2s.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub


Private Sub cmdCompute_Click()
Dim val1 As Single, val2 As Single, val3 As Single, val4 As Single, val5 As Single, val6 As Single
Dim Total As Single
Dim Gpa As Single

If Combo1.Text = "A" Then
grade1 = 4
ElseIf Combo1.Text = "AB" Then
grade1 = 3.5
ElseIf Combo1.Text = "B" Then
grade1 = 3
ElseIf Combo1.Text = "BC" Then
grade1 = 2.5
ElseIf Combo1.Text = "C" Then
grade1 = 2
ElseIf Combo1.Text = "CD" Then
grade1 = 1.5
ElseIf Combo1.Text = "D" Then
grade1 = 1
ElseIf Combo1.Text = "E" Then
grade1 = 0.5
ElseIf Combo1.Text = "F" Then
grade1 = 0
End If
If Combo2.Text = "A" Then
grade2 = 4
ElseIf Combo2.Text = "AB" Then
grade2 = 3.5
ElseIf Combo2.Text = "B" Then
grade2 = 3
ElseIf Combo2.Text = "BC" Then
grade2 = 2.5
ElseIf Combo2.Text = "C" Then
grade2 = 2
ElseIf Combo2.Text = "CD" Then
grade2 = 1.5
ElseIf Combo2.Text = "D" Then
grade2 = 1
ElseIf Combo2.Text = "E" Then
grade2 = 0.5
ElseIf Combo2.Text = "F" Then
grade2 = 0
End If
If Combo3.Text = "A" Then
grade3 = 4
ElseIf Combo3.Text = "AB" Then
grade3 = 3.5
ElseIf Combo3.Text = "B" Then
grade3 = 3
ElseIf Combo3.Text = "BC" Then
grade3 = 2.5
ElseIf Combo3.Text = "C" Then
grade3 = 2
ElseIf Combo3.Text = "CD" Then
grade3 = 1.5
ElseIf Combo3.Text = "D" Then
grade3 = 1
ElseIf Combo3.Text = "E" Then
grade3 = 0.5
ElseIf Combo3.Text = "F" Then
grade3 = 0
End If
If Combo4.Text = "A" Then
grade4 = 4
ElseIf Combo4.Text = "AB" Then
grade4 = 3.5
ElseIf Combo4.Text = "B" Then
grade4 = 3
ElseIf Combo4.Text = "BC" Then
grade4 = 2.5
ElseIf Combo4.Text = "C" Then
grade4 = 2
ElseIf Combo4.Text = "CD" Then
grade4 = 1.5
ElseIf Combo4.Text = "D" Then
grade4 = 1
ElseIf Combo4.Text = "E" Then
grade4 = 0.5
ElseIf Combo4.Text = "F" Then
grade4 = 0
End If
If Combo5.Text = "A" Then
grade5 = 4
ElseIf Combo5.Text = "AB" Then
grade5 = 3.5
ElseIf Combo5.Text = "B" Then
grade5 = 3
ElseIf Combo5.Text = "BC" Then
grade5 = 2.5
ElseIf Combo5.Text = "C" Then
grade5 = 2
ElseIf Combo5.Text = "CD" Then
grade5 = 1.5
ElseIf Combo5.Text = "D" Then
grade5 = 1
ElseIf Combo5.Text = "E" Then
grade5 = 0.5
ElseIf Combo5.Text = "F" Then
grade5 = 0
End If
If Combo6.Text = "A" Then
grade6 = 4
ElseIf Combo6.Text = "AB" Then
grade6 = 3.5
ElseIf Combo6.Text = "B" Then
grade6 = 3
ElseIf Combo6.Text = "BC" Then
grade6 = 2.5
ElseIf Combo6.Text = "C" Then
grade6 = 2
ElseIf Combo6.Text = "CD" Then
grade6 = 1.5
ElseIf Combo6.Text = "D" Then
grade6 = 1
ElseIf Combo6.Text = "E" Then
grade6 = 0.5
ElseIf Combo6.Text = "F" Then
grade6 = 0
End If


val1 = grade1 * 5
val2 = grade2 * 5
val3 = grade3 * 5
val4 = grade4 * 5
val5 = grade5 * 5
val6 = grade6 * 5


Total = val1 + val2 + val3 + val4 + val5 + val6
lblTotal.Caption = Total

Gpa = Total / totalchr
lblGpa.Caption = Gpa
End Sub


Private Sub cmdCompute2_Click()
Dim val7 As Single, val8 As Single, val9 As Single, val10 As Single, val11 As Single, val12 As Single
Dim Total2 As Single
Dim Gpa2 As Single

If Combo7.Text = "A" Then
grade7 = 4
ElseIf Combo7.Text = "AB" Then
grade7 = 3.5
ElseIf Combo7.Text = "B" Then
grade7 = 3
ElseIf Combo7.Text = "BC" Then
grade7 = 2.5
ElseIf Combo7.Text = "C" Then
grade7 = 2
ElseIf Combo7.Text = "CD" Then
grade7 = 1.5
ElseIf Combo7.Text = "D" Then
grade7 = 1
ElseIf Combo7.Text = "E" Then
grade7 = 0.5
ElseIf Combo7.Text = "F" Then
grade7 = 0
End If
If Combo8.Text = "A" Then
grade8 = 4
ElseIf Combo8.Text = "AB" Then
grade8 = 3.5
ElseIf Combo8.Text = "B" Then
grade8 = 3
ElseIf Combo8.Text = "BC" Then
grade8 = 2.5
ElseIf Combo8.Text = "C" Then
grade8 = 2
ElseIf Combo8.Text = "CD" Then
grade8 = 1.5
ElseIf Combo8.Text = "D" Then
grade8 = 1
ElseIf Combo8.Text = "E" Then
grade8 = 0.5
ElseIf Combo8.Text = "F" Then
grade8 = 0
End If
If Combo9.Text = "A" Then
grade9 = 4
ElseIf Combo9.Text = "AB" Then
grade9 = 3.5
ElseIf Combo9.Text = "B" Then
grade9 = 3
ElseIf Combo9.Text = "BC" Then
grade9 = 2.5
ElseIf Combo9.Text = "C" Then
grade9 = 2
ElseIf Combo9.Text = "CD" Then
grade9 = 1.5
ElseIf Combo9.Text = "D" Then
grade9 = 1
ElseIf Combo9.Text = "E" Then
grade9 = 0.5
ElseIf Combo9.Text = "F" Then
grade9 = 0
End If
If Combo10.Text = "A" Then
grade10 = 4
ElseIf Combo10.Text = "AB" Then
grade10 = 3.5
ElseIf Combo10.Text = "B" Then
grade10 = 3
ElseIf Combo10.Text = "BC" Then
grade10 = 2.5
ElseIf Combo10.Text = "C" Then
grade10 = 2
ElseIf Combo10.Text = "CD" Then
grade10 = 1.5
ElseIf Combo10.Text = "D" Then
grade10 = 1
ElseIf Combo10.Text = "E" Then
grade10 = 0.5
ElseIf Combo10.Text = "F" Then
grade10 = 0
End If
If Combo11.Text = "A" Then
grade11 = 4
ElseIf Combo11.Text = "AB" Then
grade11 = 3.5
ElseIf Combo11.Text = "B" Then
grade11 = 3
ElseIf Combo11.Text = "BC" Then
grade11 = 2.5
ElseIf Combo11.Text = "C" Then
grade11 = 2
ElseIf Combo11.Text = "CD" Then
grade11 = 1.5
ElseIf Combo11.Text = "D" Then
grade11 = 1
ElseIf Combo11.Text = "E" Then
grade11 = 0.5
ElseIf Combo11.Text = "F" Then
grade11 = 0
End If
If Combo12.Text = "A" Then
grade12 = 4
ElseIf Combo12.Text = "AB" Then
grade12 = 3.5
ElseIf Combo12.Text = "B" Then
grade12 = 3
ElseIf Combo12.Text = "BC" Then
grade12 = 2.5
ElseIf Combo12.Text = "C" Then
grade12 = 2
ElseIf Combo12.Text = "CD" Then
grade12 = 1.5
ElseIf Combo12.Text = "D" Then
grade12 = 1
ElseIf Combo12.Text = "E" Then
grade12 = 0.5
ElseIf Combo12.Text = "F" Then
grade12 = 0
End If


val7 = grade7 * 5
val8 = grade8 * 5
val9 = grade9 * 4
val10 = grade10 * 2
val11 = grade11 * 6
val12 = grade12 * 3


Total2 = val7 + val8 + val9 + val10 + val11 + val12
lblTotal2.Caption = Total2

Gpa2 = Total2 / totalchr2
lblGpa2.Caption = Gpa2
End Sub


Private Sub cmdSearch2_Click()
On Error GoTo joe
GetConnect2
Dim con As String
con = InputBox("Enter Student Reg. Number", "Search By Reg. No.")
BookMark1 = adoHND2EExam2s.Recordset.Bookmark
adoHND2EExam2s.Recordset.MoveFirst
adoHND2EExam2s.Recordset.Find "regno = '" & con & "'", 0, adSearchForward
If adoHND2EExam2s.Recordset.EOF = True Then
adoHND2EExam2s.Recordset.Bookmark = BookMark1
MsgBox ("No Record Found")
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub
