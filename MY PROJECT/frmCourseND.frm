VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCourseND 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Courses and Their Lecturers"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12390
   Icon            =   "frmCourseND.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "ND1(M)1S"
      TabPicture(0)   =   "frmCourseND.frx":234CD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Adodc1"
      Tab(0).Control(1)=   "cmdBack"
      Tab(0).Control(2)=   "cmdEdit"
      Tab(0).Control(3)=   "List13"
      Tab(0).Control(4)=   "List1"
      Tab(0).Control(5)=   "List2"
      Tab(0).Control(6)=   "List3"
      Tab(0).Control(7)=   "List4"
      Tab(0).Control(8)=   "List5"
      Tab(0).Control(9)=   "List6"
      Tab(0).Control(10)=   "List7"
      Tab(0).Control(11)=   "List8"
      Tab(0).Control(12)=   "List9"
      Tab(0).Control(13)=   "List10"
      Tab(0).Control(14)=   "List11"
      Tab(0).Control(15)=   "List12"
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(17)=   "Label8"
      Tab(0).Control(18)=   "Label7"
      Tab(0).Control(19)=   "Label1"
      Tab(0).Control(20)=   "Label2"
      Tab(0).Control(21)=   "Label3"
      Tab(0).Control(22)=   "Label4"
      Tab(0).Control(23)=   "Label5"
      Tab(0).Control(24)=   "Label6"
      Tab(0).Control(25)=   "A"
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "ND1(M)2S"
      TabPicture(1)   =   "frmCourseND.frx":234E9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBack1"
      Tab(1).Control(1)=   "cmdEdit1"
      Tab(1).Control(2)=   "List14"
      Tab(1).Control(3)=   "List15"
      Tab(1).Control(4)=   "List16"
      Tab(1).Control(5)=   "List17"
      Tab(1).Control(6)=   "List18"
      Tab(1).Control(7)=   "List19"
      Tab(1).Control(8)=   "List20"
      Tab(1).Control(9)=   "List21"
      Tab(1).Control(10)=   "List22"
      Tab(1).Control(11)=   "List23"
      Tab(1).Control(12)=   "List24"
      Tab(1).Control(13)=   "List25"
      Tab(1).Control(14)=   "List26"
      Tab(1).Control(15)=   "Adodc2"
      Tab(1).Control(16)=   "Label11"
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(18)=   "Label13"
      Tab(1).Control(19)=   "Label14"
      Tab(1).Control(20)=   "Label15"
      Tab(1).Control(21)=   "Label16"
      Tab(1).Control(22)=   "Label17"
      Tab(1).Control(23)=   "Label18"
      Tab(1).Control(24)=   "Label19"
      Tab(1).Control(25)=   "B"
      Tab(1).Control(26)=   "Label10"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "ND1(E)1S"
      TabPicture(2)   =   "frmCourseND.frx":23505
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBack2"
      Tab(2).Control(1)=   "cmdEdit2"
      Tab(2).Control(2)=   "List27"
      Tab(2).Control(3)=   "List28"
      Tab(2).Control(4)=   "List29"
      Tab(2).Control(5)=   "List30"
      Tab(2).Control(6)=   "List31"
      Tab(2).Control(7)=   "List32"
      Tab(2).Control(8)=   "List33"
      Tab(2).Control(9)=   "List34"
      Tab(2).Control(10)=   "List35"
      Tab(2).Control(11)=   "List36"
      Tab(2).Control(12)=   "List37"
      Tab(2).Control(13)=   "List38"
      Tab(2).Control(14)=   "List39"
      Tab(2).Control(15)=   "Adodc3"
      Tab(2).Control(16)=   "Label20"
      Tab(2).Control(17)=   "Label21"
      Tab(2).Control(18)=   "Label22"
      Tab(2).Control(19)=   "Label23"
      Tab(2).Control(20)=   "Label24"
      Tab(2).Control(21)=   "Label25"
      Tab(2).Control(22)=   "Label26"
      Tab(2).Control(23)=   "Label27"
      Tab(2).Control(24)=   "Label28"
      Tab(2).Control(25)=   "Label29"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "ND1(E)2S"
      TabPicture(3)   =   "frmCourseND.frx":23521
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdBack3"
      Tab(3).Control(1)=   "cmdEdit3"
      Tab(3).Control(2)=   "List40"
      Tab(3).Control(3)=   "List41"
      Tab(3).Control(4)=   "List42"
      Tab(3).Control(5)=   "List43"
      Tab(3).Control(6)=   "List44"
      Tab(3).Control(7)=   "List45"
      Tab(3).Control(8)=   "List46"
      Tab(3).Control(9)=   "List47"
      Tab(3).Control(10)=   "List48"
      Tab(3).Control(11)=   "List49"
      Tab(3).Control(12)=   "List50"
      Tab(3).Control(13)=   "List51"
      Tab(3).Control(14)=   "List52"
      Tab(3).Control(15)=   "Adodc4"
      Tab(3).Control(16)=   "Label70"
      Tab(3).Control(17)=   "Label30"
      Tab(3).Control(18)=   "Label31"
      Tab(3).Control(19)=   "Label32"
      Tab(3).Control(20)=   "Label33"
      Tab(3).Control(21)=   "Label34"
      Tab(3).Control(22)=   "Label35"
      Tab(3).Control(23)=   "Label36"
      Tab(3).Control(24)=   "Label37"
      Tab(3).Control(25)=   "Label38"
      Tab(3).Control(26)=   "Label39"
      Tab(3).ControlCount=   27
      TabCaption(4)   =   "ND2(M)1S"
      TabPicture(4)   =   "frmCourseND.frx":2353D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdBack4"
      Tab(4).Control(1)=   "cmdEdit4"
      Tab(4).Control(2)=   "List53"
      Tab(4).Control(3)=   "List54"
      Tab(4).Control(4)=   "List55"
      Tab(4).Control(5)=   "List56"
      Tab(4).Control(6)=   "List57"
      Tab(4).Control(7)=   "List58"
      Tab(4).Control(8)=   "List59"
      Tab(4).Control(9)=   "List60"
      Tab(4).Control(10)=   "List61"
      Tab(4).Control(11)=   "List62"
      Tab(4).Control(12)=   "List63"
      Tab(4).Control(13)=   "List64"
      Tab(4).Control(14)=   "List65"
      Tab(4).Control(15)=   "Adodc5"
      Tab(4).Control(16)=   "Label71"
      Tab(4).Control(17)=   "Label40"
      Tab(4).Control(18)=   "Label41"
      Tab(4).Control(19)=   "Label42"
      Tab(4).Control(20)=   "Label43"
      Tab(4).Control(21)=   "Label44"
      Tab(4).Control(22)=   "Label45"
      Tab(4).Control(23)=   "Label46"
      Tab(4).ControlCount=   24
      TabCaption(5)   =   "ND2(M)2S"
      TabPicture(5)   =   "frmCourseND.frx":23559
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdBack5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdEdit5"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "List66"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "List67"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "List68"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "List69"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "List70"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "List71"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "List72"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "List73"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "List74"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "List75"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "List76"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "List77"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "List78"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Adodc6"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Label72"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Label47"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Label48"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "Label49"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Label50"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "Label51"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Label52"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "Label53"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "Label54"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).ControlCount=   25
      TabCaption(6)   =   "ND2(E)1S"
      TabPicture(6)   =   "frmCourseND.frx":23575
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label61"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label60"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label59"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label58"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Label57"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Label56"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Label55"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "Label73"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Adodc7"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "List91"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "List90"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "List89"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "List88"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "List87"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "List86"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).Control(15)=   "List85"
      Tab(6).Control(15).Enabled=   0   'False
      Tab(6).Control(16)=   "List84"
      Tab(6).Control(16).Enabled=   0   'False
      Tab(6).Control(17)=   "List83"
      Tab(6).Control(17).Enabled=   0   'False
      Tab(6).Control(18)=   "List82"
      Tab(6).Control(18).Enabled=   0   'False
      Tab(6).Control(19)=   "List81"
      Tab(6).Control(19).Enabled=   0   'False
      Tab(6).Control(20)=   "List80"
      Tab(6).Control(20).Enabled=   0   'False
      Tab(6).Control(21)=   "List79"
      Tab(6).Control(21).Enabled=   0   'False
      Tab(6).Control(22)=   "cmdEdit6"
      Tab(6).Control(22).Enabled=   0   'False
      Tab(6).Control(23)=   "cmdBack6"
      Tab(6).Control(23).Enabled=   0   'False
      Tab(6).ControlCount=   24
      TabCaption(7)   =   "ND2(E)2S"
      TabPicture(7)   =   "frmCourseND.frx":23591
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdBack7"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "cmdEdit7"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "List92"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "List93"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "List94"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "List95"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "List96"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "List97"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "List98"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "List99"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "List100"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "List101"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).Control(12)=   "List102"
      Tab(7).Control(12).Enabled=   0   'False
      Tab(7).Control(13)=   "List103"
      Tab(7).Control(13).Enabled=   0   'False
      Tab(7).Control(14)=   "List104"
      Tab(7).Control(14).Enabled=   0   'False
      Tab(7).Control(15)=   "Adodc8"
      Tab(7).Control(15).Enabled=   0   'False
      Tab(7).Control(16)=   "Label74"
      Tab(7).Control(16).Enabled=   0   'False
      Tab(7).Control(17)=   "Label62"
      Tab(7).Control(17).Enabled=   0   'False
      Tab(7).Control(18)=   "Label63"
      Tab(7).Control(18).Enabled=   0   'False
      Tab(7).Control(19)=   "Label64"
      Tab(7).Control(19).Enabled=   0   'False
      Tab(7).Control(20)=   "Label65"
      Tab(7).Control(20).Enabled=   0   'False
      Tab(7).Control(21)=   "Label66"
      Tab(7).Control(21).Enabled=   0   'False
      Tab(7).Control(22)=   "Label67"
      Tab(7).Control(22).Enabled=   0   'False
      Tab(7).Control(23)=   "Label68"
      Tab(7).Control(23).Enabled=   0   'False
      Tab(7).Control(24)=   "Label69"
      Tab(7).Control(24).Enabled=   0   'False
      Tab(7).ControlCount=   25
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   -75000
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND1M1stS"
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
      Begin VB.CommandButton cmdBack7 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   197
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit7 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   196
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack6 
         Caption         =   "&Back"
         Height          =   495
         Left            =   6000
         TabIndex        =   195
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit6 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   4560
         TabIndex        =   194
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack5 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69000
         TabIndex        =   193
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit5 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70440
         TabIndex        =   192
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack4 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69000
         TabIndex        =   191
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit4 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70440
         TabIndex        =   190
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack3 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69000
         TabIndex        =   189
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit3 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70440
         TabIndex        =   188
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack2 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   187
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit2 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   186
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack1 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -68880
         TabIndex        =   185
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit1 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70320
         TabIndex        =   184
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69240
         TabIndex        =   183
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70680
         TabIndex        =   182
         Top             =   5520
         Width           =   1335
      End
      Begin VB.ListBox List13 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":235AD
         Left            =   -66120
         List            =   "frmCourseND.frx":235B4
         TabIndex        =   105
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":235C3
         Left            =   -74760
         List            =   "frmCourseND.frx":235E2
         TabIndex        =   104
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23630
         Left            =   -73320
         List            =   "frmCourseND.frx":2364F
         TabIndex        =   103
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":2374E
         Left            =   -68280
         List            =   "frmCourseND.frx":2376D
         TabIndex        =   102
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":2378C
         Left            =   -67920
         List            =   "frmCourseND.frx":237AB
         TabIndex        =   101
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":237CA
         Left            =   -67560
         List            =   "frmCourseND.frx":237EC
         TabIndex        =   100
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":2380F
         Left            =   -66960
         List            =   "frmCourseND.frx":2382E
         TabIndex        =   99
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2385C
         Left            =   -74760
         List            =   "frmCourseND.frx":23863
         TabIndex        =   98
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23870
         Left            =   -73320
         List            =   "frmCourseND.frx":23877
         TabIndex        =   97
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23889
         Left            =   -68280
         List            =   "frmCourseND.frx":23890
         TabIndex        =   96
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23897
         Left            =   -67920
         List            =   "frmCourseND.frx":2389E
         TabIndex        =   95
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":238A5
         Left            =   -67560
         List            =   "frmCourseND.frx":238AC
         TabIndex        =   94
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":238B4
         Left            =   -66960
         List            =   "frmCourseND.frx":238BB
         TabIndex        =   93
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List14 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":238C5
         Left            =   -66120
         List            =   "frmCourseND.frx":238CC
         TabIndex        =   92
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List15 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":238DB
         Left            =   -74760
         List            =   "frmCourseND.frx":238FD
         TabIndex        =   91
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List16 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23951
         Left            =   -73320
         List            =   "frmCourseND.frx":23973
         TabIndex        =   90
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List17 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23AA3
         Left            =   -68280
         List            =   "frmCourseND.frx":23AC5
         TabIndex        =   89
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List18 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23AE7
         Left            =   -67920
         List            =   "frmCourseND.frx":23B09
         TabIndex        =   88
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List19 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23B2B
         Left            =   -67560
         List            =   "frmCourseND.frx":23B50
         TabIndex        =   87
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List20 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23B76
         Left            =   -66960
         List            =   "frmCourseND.frx":23B98
         TabIndex        =   86
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List21 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23BC8
         Left            =   -74760
         List            =   "frmCourseND.frx":23BCF
         TabIndex        =   85
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List22 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23BDC
         Left            =   -73320
         List            =   "frmCourseND.frx":23BE3
         TabIndex        =   84
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List23 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23BF5
         Left            =   -68280
         List            =   "frmCourseND.frx":23BFC
         TabIndex        =   83
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List24 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C03
         Left            =   -67920
         List            =   "frmCourseND.frx":23C0A
         TabIndex        =   82
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List25 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C11
         Left            =   -67560
         List            =   "frmCourseND.frx":23C18
         TabIndex        =   81
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List26 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C20
         Left            =   -66960
         List            =   "frmCourseND.frx":23C27
         TabIndex        =   80
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List27 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C31
         Left            =   -66960
         List            =   "frmCourseND.frx":23C38
         TabIndex        =   79
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List28 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C42
         Left            =   -67560
         List            =   "frmCourseND.frx":23C49
         TabIndex        =   78
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List29 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C51
         Left            =   -67920
         List            =   "frmCourseND.frx":23C58
         TabIndex        =   77
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List30 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C5F
         Left            =   -68280
         List            =   "frmCourseND.frx":23C66
         TabIndex        =   76
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List31 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C6D
         Left            =   -73320
         List            =   "frmCourseND.frx":23C74
         TabIndex        =   75
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List32 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23C86
         Left            =   -74760
         List            =   "frmCourseND.frx":23C8D
         TabIndex        =   74
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List33 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23C9A
         Left            =   -66960
         List            =   "frmCourseND.frx":23CB9
         TabIndex        =   73
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List34 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23CE7
         Left            =   -67560
         List            =   "frmCourseND.frx":23D09
         TabIndex        =   72
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List35 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23D2C
         Left            =   -67920
         List            =   "frmCourseND.frx":23D4B
         TabIndex        =   71
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List36 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23D6A
         Left            =   -68280
         List            =   "frmCourseND.frx":23D89
         TabIndex        =   70
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List37 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23DA8
         Left            =   -73320
         List            =   "frmCourseND.frx":23DC7
         TabIndex        =   69
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List38 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3990
         ItemData        =   "frmCourseND.frx":23EC6
         Left            =   -74760
         List            =   "frmCourseND.frx":23EE5
         TabIndex        =   68
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List39 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F33
         Left            =   -66120
         List            =   "frmCourseND.frx":23F3A
         TabIndex        =   67
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List40 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F49
         Left            =   -66960
         List            =   "frmCourseND.frx":23F50
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List41 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F5A
         Left            =   -67560
         List            =   "frmCourseND.frx":23F61
         TabIndex        =   65
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List42 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F69
         Left            =   -67920
         List            =   "frmCourseND.frx":23F70
         TabIndex        =   64
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List43 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F77
         Left            =   -68280
         List            =   "frmCourseND.frx":23F7E
         TabIndex        =   63
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List44 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F85
         Left            =   -73320
         List            =   "frmCourseND.frx":23F8C
         TabIndex        =   62
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List45 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":23F9E
         Left            =   -74760
         List            =   "frmCourseND.frx":23FA5
         TabIndex        =   61
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List46 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":23FB2
         Left            =   -66960
         List            =   "frmCourseND.frx":23FD4
         TabIndex        =   60
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List47 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":24004
         Left            =   -67560
         List            =   "frmCourseND.frx":24029
         TabIndex        =   59
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List48 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":2404F
         Left            =   -67920
         List            =   "frmCourseND.frx":24071
         TabIndex        =   58
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List49 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":24093
         Left            =   -68280
         List            =   "frmCourseND.frx":240B5
         TabIndex        =   57
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List50 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":240D7
         Left            =   -73320
         List            =   "frmCourseND.frx":240F9
         TabIndex        =   56
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List51 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmCourseND.frx":24229
         Left            =   -74760
         List            =   "frmCourseND.frx":2424B
         TabIndex        =   55
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List52 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2429F
         Left            =   -66120
         List            =   "frmCourseND.frx":242A6
         TabIndex        =   54
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List53 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":242B5
         Left            =   -66960
         List            =   "frmCourseND.frx":242BC
         TabIndex        =   53
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List54 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":242C6
         Left            =   -67560
         List            =   "frmCourseND.frx":242CD
         TabIndex        =   52
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List55 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":242D5
         Left            =   -67920
         List            =   "frmCourseND.frx":242DC
         TabIndex        =   51
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List56 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":242E3
         Left            =   -68280
         List            =   "frmCourseND.frx":242EA
         TabIndex        =   50
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List57 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":242F1
         Left            =   -73320
         List            =   "frmCourseND.frx":242F8
         TabIndex        =   49
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List58 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2430A
         Left            =   -74760
         List            =   "frmCourseND.frx":24311
         TabIndex        =   48
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List59 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":2431E
         Left            =   -66960
         List            =   "frmCourseND.frx":24337
         TabIndex        =   47
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List60 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24358
         Left            =   -67560
         List            =   "frmCourseND.frx":24374
         TabIndex        =   46
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List61 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24391
         Left            =   -67920
         List            =   "frmCourseND.frx":243AA
         TabIndex        =   45
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List62 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":243C3
         Left            =   -68280
         List            =   "frmCourseND.frx":243DC
         TabIndex        =   44
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List63 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":243F5
         Left            =   -73320
         List            =   "frmCourseND.frx":2440E
         TabIndex        =   43
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List64 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":244D3
         Left            =   -74760
         List            =   "frmCourseND.frx":244EC
         TabIndex        =   42
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List65 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24528
         Left            =   -66120
         List            =   "frmCourseND.frx":2452F
         TabIndex        =   41
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List66 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2453E
         Left            =   -66960
         List            =   "frmCourseND.frx":24545
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List67 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2454F
         Left            =   -67560
         List            =   "frmCourseND.frx":24556
         TabIndex        =   39
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List68 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2455E
         Left            =   -67920
         List            =   "frmCourseND.frx":24565
         TabIndex        =   38
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List69 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2456C
         Left            =   -68280
         List            =   "frmCourseND.frx":24573
         TabIndex        =   37
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List70 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2457A
         Left            =   -73320
         List            =   "frmCourseND.frx":24581
         TabIndex        =   36
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List71 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24593
         Left            =   -74760
         List            =   "frmCourseND.frx":2459A
         TabIndex        =   35
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List72 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":245A7
         Left            =   -66960
         List            =   "frmCourseND.frx":245C3
         TabIndex        =   34
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List73 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":245E8
         Left            =   -67560
         List            =   "frmCourseND.frx":24607
         TabIndex        =   33
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List74 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24627
         Left            =   -67920
         List            =   "frmCourseND.frx":24643
         TabIndex        =   32
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List75 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":2465F
         Left            =   -68280
         List            =   "frmCourseND.frx":2467B
         TabIndex        =   31
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List76 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24697
         Left            =   -73320
         List            =   "frmCourseND.frx":246B3
         TabIndex        =   30
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List77 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24780
         Left            =   -74760
         List            =   "frmCourseND.frx":2479C
         TabIndex        =   29
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List78 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":247E0
         Left            =   -66120
         List            =   "frmCourseND.frx":247E7
         TabIndex        =   28
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List79 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":247F6
         Left            =   8040
         List            =   "frmCourseND.frx":247FD
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List80 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24807
         Left            =   7440
         List            =   "frmCourseND.frx":2480E
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List81 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24816
         Left            =   7080
         List            =   "frmCourseND.frx":2481D
         TabIndex        =   25
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List82 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24824
         Left            =   6720
         List            =   "frmCourseND.frx":2482B
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List83 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24832
         Left            =   1680
         List            =   "frmCourseND.frx":24839
         TabIndex        =   23
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List84 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":2484B
         Left            =   240
         List            =   "frmCourseND.frx":24852
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List85 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":2485F
         Left            =   8040
         List            =   "frmCourseND.frx":24878
         TabIndex        =   21
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List86 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24899
         Left            =   7440
         List            =   "frmCourseND.frx":248B5
         TabIndex        =   20
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List87 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":248D2
         Left            =   7080
         List            =   "frmCourseND.frx":248EB
         TabIndex        =   19
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List88 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24904
         Left            =   6720
         List            =   "frmCourseND.frx":2491D
         TabIndex        =   18
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List89 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24936
         Left            =   1680
         List            =   "frmCourseND.frx":2494F
         TabIndex        =   17
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List90 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmCourseND.frx":24A14
         Left            =   240
         List            =   "frmCourseND.frx":24A2D
         TabIndex        =   16
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List91 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24A69
         Left            =   8880
         List            =   "frmCourseND.frx":24A70
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List92 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24A7F
         Left            =   -66960
         List            =   "frmCourseND.frx":24A86
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List93 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24A90
         Left            =   -67560
         List            =   "frmCourseND.frx":24A97
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List94 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24A9F
         Left            =   -67920
         List            =   "frmCourseND.frx":24AA6
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List95 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24AAD
         Left            =   -68280
         List            =   "frmCourseND.frx":24AB4
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List96 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24ABB
         Left            =   -73320
         List            =   "frmCourseND.frx":24AC2
         TabIndex        =   10
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List97 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24AD4
         Left            =   -74760
         List            =   "frmCourseND.frx":24ADB
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List98 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24AE8
         Left            =   -66960
         List            =   "frmCourseND.frx":24B04
         TabIndex        =   8
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List99 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24B29
         Left            =   -67560
         List            =   "frmCourseND.frx":24B48
         TabIndex        =   7
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List100 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24B68
         Left            =   -67920
         List            =   "frmCourseND.frx":24B84
         TabIndex        =   6
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List101 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24BA0
         Left            =   -68280
         List            =   "frmCourseND.frx":24BBC
         TabIndex        =   5
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List102 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24BD8
         Left            =   -73320
         List            =   "frmCourseND.frx":24BF4
         TabIndex        =   4
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List103 
         Appearance      =   0  'Flat
         Columns         =   1
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "TimelessTLig"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "frmCourseND.frx":24CC1
         Left            =   -74760
         List            =   "frmCourseND.frx":24CDD
         TabIndex        =   3
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List104 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCourseND.frx":24D21
         Left            =   -66120
         List            =   "frmCourseND.frx":24D28
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND1M2ndS"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND1E1stS"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND1E2ndS"
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
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND2M1stS"
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
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND2M2ndS"
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
      Begin MSAdodcLib.Adodc Adodc7 
         Height          =   735
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND2E1stS"
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
      Begin MSAdodcLib.Adodc Adodc8 
         Height          =   735
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=LectureSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "LectureSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblND2E2ndS"
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
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "ND II(EVENING) SECOND SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   181
         Top             =   480
         Width           =   5730
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "ND II(EVENING) FIRST SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3240
         TabIndex        =   180
         Top             =   480
         Width           =   5430
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "ND II(MORNING) SECOND SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   179
         Top             =   480
         Width           =   5805
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "ND II(MORNING) FIRST SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   178
         Top             =   480
         Width           =   5505
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "ND I(EVENING) SECOND SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   177
         Top             =   480
         Width           =   5625
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS127"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   176
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   175
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MATH112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   174
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM101"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   173
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   172
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM113"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   171
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA111"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   170
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   169
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MATH111"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   168
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM122"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   167
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM123"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   166
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM124"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   165
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM125"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   164
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM126"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   163
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS102"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   162
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS128"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   161
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         DataField       =   "EED126"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   160
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         DataField       =   "URP120"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   159
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label A 
         AutoSize        =   -1  'True
         Caption         =   "ND I(MORNING) FIRST SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   158
         Top             =   480
         Width           =   5400
      End
      Begin VB.Label B 
         AutoSize        =   -1  'True
         Caption         =   "ND I(MORNING) SECOND SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   157
         Top             =   480
         Width           =   5700
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM121"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   156
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM101"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   155
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   154
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM113"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   153
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA111"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   152
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   151
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MATH111"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   150
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MATH112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   149
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   148
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS127"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   147
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "ND I(EVENING) FIRST SEMESTER"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   146
         Top             =   480
         Width           =   5325
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM121"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   145
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM122"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   144
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM123"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   143
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM124"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   142
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM125"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   141
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM126"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   140
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS102"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   139
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS128"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   138
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         DataField       =   "EED126"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   137
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         DataField       =   "URP120"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   136
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM211"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   135
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM212"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   134
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM213"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   133
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM214"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   132
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM215"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   131
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM216"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   130
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS201"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   129
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM221"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   128
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM222"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   127
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM223"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   126
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM224"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   125
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM225"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   124
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM226"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   123
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label53 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA226"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   122
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM229"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   121
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label55 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM211"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   120
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label56 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM212"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   119
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label57 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM213"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   118
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label58 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM214"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   117
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label59 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM215"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   116
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label60 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM216"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   115
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label61 
         BackColor       =   &H00FFFFFF&
         DataField       =   "GNS201"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   114
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label62 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM221"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   113
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label63 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM222"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   112
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label64 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM223"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   111
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label65 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM224"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   110
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label66 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM225"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   109
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label67 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM226"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   108
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label68 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA226"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   107
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label69 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM229"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "GoudyHandtooled BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66120
         TabIndex        =   106
         Top             =   3840
         Width           =   2895
      End
   End
   Begin VB.Label logo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTAL COURSES for ND"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3120
      TabIndex        =   0
      Top             =   2880
      Width           =   5985
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmCourseND.frx":24D37
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmCourseND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect1()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect2()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect3()
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect4()
Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect5()
Adodc5.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect6()
Adodc6.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect7()
Adodc7.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Function GetConnect8()
Adodc8.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function

Private Sub cmdBack_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack1_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdEdit_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 0
End Sub
Private Sub cmdBack2_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack3_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack4_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack5_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack6_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdBack7_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdEdit1_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 1
End Sub

Private Sub cmdEdit2_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 2
End Sub

Private Sub cmdEdit3_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 3
End Sub

Private Sub cmdEdit4_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 4
End Sub

Private Sub cmdEdit5_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 5
End Sub

Private Sub cmdEdit6_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 6
End Sub

Private Sub cmdEdit7_Click()
Me.Hide
frmCourseNDEdit.Show
frmCourseNDEdit.SSTab1.Tab = 7
End Sub

Private Sub Form_Load()
GetConnect1
GetConnect2
GetConnect3
GetConnect4
GetConnect5
GetConnect6
GetConnect7
GetConnect8
End Sub
