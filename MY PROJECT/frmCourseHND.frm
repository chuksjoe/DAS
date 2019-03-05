VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCourseHND 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Courses and their Lecturers"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12390
   Icon            =   "frmCourseHND.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "HND I 1st Sem"
      TabPicture(0)   =   "frmCourseHND.frx":234CD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "A"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Adodc1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List117"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "List116"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List115"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "List114"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List113"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "List112"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "List111"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "List110"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "List109"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "List108"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "List107"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "List106"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "List105"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdEdit"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdBack"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "HND I 2nd Sem"
      TabPicture(1)   =   "frmCourseHND.frx":234E9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "B"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Adodc2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "List130"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "List129"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "List128"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "List127"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "List126"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "List125"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "List124"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "List123"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "List122"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "List121"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "List120"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "List119"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "List118"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdEdit1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdBack1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "HND II 1st Sem"
      TabPicture(2)   =   "frmCourseHND.frx":23505
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label19"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(5)=   "Label15"
      Tab(2).Control(6)=   "C"
      Tab(2).Control(7)=   "Adodc3"
      Tab(2).Control(8)=   "List143"
      Tab(2).Control(9)=   "List142"
      Tab(2).Control(10)=   "List141"
      Tab(2).Control(11)=   "List140"
      Tab(2).Control(12)=   "List139"
      Tab(2).Control(13)=   "List138"
      Tab(2).Control(14)=   "List137"
      Tab(2).Control(15)=   "List136"
      Tab(2).Control(16)=   "List135"
      Tab(2).Control(17)=   "List134"
      Tab(2).Control(18)=   "List133"
      Tab(2).Control(19)=   "List132"
      Tab(2).Control(20)=   "List131"
      Tab(2).Control(21)=   "cmdEdit2"
      Tab(2).Control(22)=   "cmdBack2"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "HND II 2nd Sem"
      TabPicture(3)   =   "frmCourseHND.frx":23521
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label23"
      Tab(3).Control(3)=   "Label24"
      Tab(3).Control(4)=   "Label25"
      Tab(3).Control(5)=   "Label26"
      Tab(3).Control(6)=   "D"
      Tab(3).Control(7)=   "Adodc4"
      Tab(3).Control(8)=   "List1"
      Tab(3).Control(9)=   "List2"
      Tab(3).Control(10)=   "List3"
      Tab(3).Control(11)=   "List4"
      Tab(3).Control(12)=   "List5"
      Tab(3).Control(13)=   "List6"
      Tab(3).Control(14)=   "List7"
      Tab(3).Control(15)=   "List8"
      Tab(3).Control(16)=   "List9"
      Tab(3).Control(17)=   "List10"
      Tab(3).Control(18)=   "List11"
      Tab(3).Control(19)=   "List12"
      Tab(3).Control(20)=   "List13"
      Tab(3).Control(21)=   "cmdEdit3"
      Tab(3).Control(22)=   "cmdBack3"
      Tab(3).ControlCount=   23
      TabCaption(4)   =   "HND1(E) 1st S"
      TabPicture(4)   =   "frmCourseHND.frx":2353D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "E"
      Tab(4).Control(1)=   "Label28"
      Tab(4).Control(2)=   "Label29"
      Tab(4).Control(3)=   "Label30"
      Tab(4).Control(4)=   "Label31"
      Tab(4).Control(5)=   "Label32"
      Tab(4).Control(6)=   "Label33"
      Tab(4).Control(7)=   "Label34"
      Tab(4).Control(8)=   "Adodc5"
      Tab(4).Control(9)=   "cmdBack4"
      Tab(4).Control(10)=   "cmdEdit4"
      Tab(4).Control(11)=   "List14"
      Tab(4).Control(12)=   "List15"
      Tab(4).Control(13)=   "List16"
      Tab(4).Control(14)=   "List17"
      Tab(4).Control(15)=   "List18"
      Tab(4).Control(16)=   "List19"
      Tab(4).Control(17)=   "List20"
      Tab(4).Control(18)=   "List21"
      Tab(4).Control(19)=   "List22"
      Tab(4).Control(20)=   "List23"
      Tab(4).Control(21)=   "List24"
      Tab(4).Control(22)=   "List25"
      Tab(4).Control(23)=   "List26"
      Tab(4).ControlCount=   24
      TabCaption(5)   =   "HND1(E) 2nd S"
      TabPicture(5)   =   "frmCourseHND.frx":23559
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label35"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label36"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label37"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label38"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label39"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label40"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label41"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "F"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Adodc6"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdBack5"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "cmdEdit5"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "List27"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "List28"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "List29"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "List30"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "List31"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "List32"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "List33"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "List34"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "List35"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "List36"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "List37"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "List38"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "List39"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).ControlCount=   24
      TabCaption(6)   =   "HND2(E) 1st S"
      TabPicture(6)   =   "frmCourseHND.frx":23575
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "List52"
      Tab(6).Control(1)=   "List51"
      Tab(6).Control(2)=   "List50"
      Tab(6).Control(3)=   "List49"
      Tab(6).Control(4)=   "List48"
      Tab(6).Control(5)=   "List47"
      Tab(6).Control(6)=   "List46"
      Tab(6).Control(7)=   "List45"
      Tab(6).Control(8)=   "List44"
      Tab(6).Control(9)=   "List43"
      Tab(6).Control(10)=   "List42"
      Tab(6).Control(11)=   "List41"
      Tab(6).Control(12)=   "List40"
      Tab(6).Control(13)=   "cmdEdit6"
      Tab(6).Control(14)=   "cmdBack6"
      Tab(6).Control(15)=   "Adodc7"
      Tab(6).Control(16)=   "Label47"
      Tab(6).Control(17)=   "Label46"
      Tab(6).Control(18)=   "Label45"
      Tab(6).Control(19)=   "Label44"
      Tab(6).Control(20)=   "Label43"
      Tab(6).Control(21)=   "Label42"
      Tab(6).Control(22)=   "G"
      Tab(6).ControlCount=   23
      TabCaption(7)   =   "HND2(E) 2nd S"
      TabPicture(7)   =   "frmCourseHND.frx":23591
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label27"
      Tab(7).Control(1)=   "Label48"
      Tab(7).Control(2)=   "Label49"
      Tab(7).Control(3)=   "Label50"
      Tab(7).Control(4)=   "Label51"
      Tab(7).Control(5)=   "Label52"
      Tab(7).Control(6)=   "Label53"
      Tab(7).Control(7)=   "Adodc8"
      Tab(7).Control(8)=   "cmdBack7"
      Tab(7).Control(9)=   "cmdEdit7"
      Tab(7).Control(10)=   "List53"
      Tab(7).Control(11)=   "List54"
      Tab(7).Control(12)=   "List55"
      Tab(7).Control(13)=   "List56"
      Tab(7).Control(14)=   "List57"
      Tab(7).Control(15)=   "List58"
      Tab(7).Control(16)=   "List59"
      Tab(7).Control(17)=   "List60"
      Tab(7).Control(18)=   "List61"
      Tab(7).Control(19)=   "List62"
      Tab(7).Control(20)=   "List63"
      Tab(7).Control(21)=   "List64"
      Tab(7).Control(22)=   "List65"
      Tab(7).ControlCount=   23
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
         ItemData        =   "frmCourseHND.frx":235AD
         Left            =   -66960
         List            =   "frmCourseHND.frx":235B4
         TabIndex        =   174
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List64 
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
         ItemData        =   "frmCourseHND.frx":235BE
         Left            =   -67560
         List            =   "frmCourseHND.frx":235C5
         TabIndex        =   173
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List63 
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
         ItemData        =   "frmCourseHND.frx":235CD
         Left            =   -67920
         List            =   "frmCourseHND.frx":235D4
         TabIndex        =   172
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List62 
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
         ItemData        =   "frmCourseHND.frx":235DB
         Left            =   -68280
         List            =   "frmCourseHND.frx":235E2
         TabIndex        =   171
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List61 
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
         ItemData        =   "frmCourseHND.frx":235E9
         Left            =   -73320
         List            =   "frmCourseHND.frx":235F0
         TabIndex        =   170
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List60 
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
         ItemData        =   "frmCourseHND.frx":23602
         Left            =   -74760
         List            =   "frmCourseHND.frx":23609
         TabIndex        =   169
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23616
         Left            =   -66960
         List            =   "frmCourseHND.frx":2362F
         TabIndex        =   168
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List58 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2364E
         Left            =   -67560
         List            =   "frmCourseHND.frx":2366A
         TabIndex        =   167
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List57 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23686
         Left            =   -67920
         List            =   "frmCourseHND.frx":2369F
         TabIndex        =   166
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List56 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":236B9
         Left            =   -68280
         List            =   "frmCourseHND.frx":236D2
         TabIndex        =   165
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List55 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":236EB
         Left            =   -73320
         List            =   "frmCourseHND.frx":23704
         TabIndex        =   164
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List54 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":237C4
         Left            =   -74760
         List            =   "frmCourseHND.frx":237DD
         TabIndex        =   163
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHND.frx":23813
         Left            =   -66120
         List            =   "frmCourseHND.frx":2381A
         TabIndex        =   162
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdEdit7 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   161
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack7 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   160
         Top             =   4800
         Width           =   1335
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
         ItemData        =   "frmCourseHND.frx":23829
         Left            =   -66120
         List            =   "frmCourseHND.frx":23830
         TabIndex        =   152
         Top             =   960
         Width           =   2895
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2383F
         Left            =   -74760
         List            =   "frmCourseHND.frx":23855
         TabIndex        =   151
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23889
         Left            =   -73320
         List            =   "frmCourseHND.frx":2389F
         TabIndex        =   150
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23943
         Left            =   -68280
         List            =   "frmCourseHND.frx":23959
         TabIndex        =   149
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2396F
         Left            =   -67920
         List            =   "frmCourseHND.frx":23985
         TabIndex        =   148
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List47 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2399B
         Left            =   -67560
         List            =   "frmCourseHND.frx":239B4
         TabIndex        =   147
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":239CE
         Left            =   -66960
         List            =   "frmCourseHND.frx":239E4
         TabIndex        =   146
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHND.frx":23A01
         Left            =   -74760
         List            =   "frmCourseHND.frx":23A08
         TabIndex        =   145
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHND.frx":23A15
         Left            =   -73320
         List            =   "frmCourseHND.frx":23A1C
         TabIndex        =   144
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHND.frx":23A2E
         Left            =   -68280
         List            =   "frmCourseHND.frx":23A35
         TabIndex        =   143
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHND.frx":23A3C
         Left            =   -67920
         List            =   "frmCourseHND.frx":23A43
         TabIndex        =   142
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHND.frx":23A4A
         Left            =   -67560
         List            =   "frmCourseHND.frx":23A51
         TabIndex        =   141
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHND.frx":23A59
         Left            =   -66960
         List            =   "frmCourseHND.frx":23A60
         TabIndex        =   140
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit6 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70200
         TabIndex        =   139
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack6 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -68760
         TabIndex        =   138
         Top             =   4800
         Width           =   1335
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
         ItemData        =   "frmCourseHND.frx":23A6A
         Left            =   -66120
         List            =   "frmCourseHND.frx":23A71
         TabIndex        =   129
         Top             =   960
         Width           =   2895
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23A80
         Left            =   -74760
         List            =   "frmCourseHND.frx":23A99
         TabIndex        =   128
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23AD5
         Left            =   -73320
         List            =   "frmCourseHND.frx":23AEE
         TabIndex        =   127
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23BA2
         Left            =   -68280
         List            =   "frmCourseHND.frx":23BBB
         TabIndex        =   126
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23BD4
         Left            =   -67920
         List            =   "frmCourseHND.frx":23BED
         TabIndex        =   125
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23C06
         Left            =   -67560
         List            =   "frmCourseHND.frx":23C22
         TabIndex        =   124
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23C3F
         Left            =   -66960
         List            =   "frmCourseHND.frx":23C58
         TabIndex        =   123
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHND.frx":23C7A
         Left            =   -74760
         List            =   "frmCourseHND.frx":23C81
         TabIndex        =   122
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHND.frx":23C8E
         Left            =   -73320
         List            =   "frmCourseHND.frx":23C95
         TabIndex        =   121
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHND.frx":23CA7
         Left            =   -68280
         List            =   "frmCourseHND.frx":23CAE
         TabIndex        =   120
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHND.frx":23CB5
         Left            =   -67920
         List            =   "frmCourseHND.frx":23CBC
         TabIndex        =   119
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHND.frx":23CC3
         Left            =   -67560
         List            =   "frmCourseHND.frx":23CCA
         TabIndex        =   118
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHND.frx":23CD2
         Left            =   -66960
         List            =   "frmCourseHND.frx":23CD9
         TabIndex        =   117
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit5 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   116
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack5 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   115
         Top             =   4920
         Width           =   1335
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
         ItemData        =   "frmCourseHND.frx":23CE3
         Left            =   -66120
         List            =   "frmCourseHND.frx":23CEA
         TabIndex        =   106
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List25 
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
         ItemData        =   "frmCourseHND.frx":23CF9
         Left            =   -74760
         List            =   "frmCourseHND.frx":23D12
         TabIndex        =   105
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List24 
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
         ItemData        =   "frmCourseHND.frx":23D4E
         Left            =   -73320
         List            =   "frmCourseHND.frx":23D67
         TabIndex        =   104
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List23 
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
         ItemData        =   "frmCourseHND.frx":23E0F
         Left            =   -68280
         List            =   "frmCourseHND.frx":23E28
         TabIndex        =   103
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List22 
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
         ItemData        =   "frmCourseHND.frx":23E41
         Left            =   -67920
         List            =   "frmCourseHND.frx":23E5A
         TabIndex        =   102
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List21 
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
         ItemData        =   "frmCourseHND.frx":23E73
         Left            =   -67560
         List            =   "frmCourseHND.frx":23E8F
         TabIndex        =   101
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
         Height          =   3270
         ItemData        =   "frmCourseHND.frx":23EAC
         Left            =   -66960
         List            =   "frmCourseHND.frx":23EC5
         TabIndex        =   100
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List19 
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
         ItemData        =   "frmCourseHND.frx":23EE8
         Left            =   -74760
         List            =   "frmCourseHND.frx":23EEF
         TabIndex        =   99
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List18 
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
         ItemData        =   "frmCourseHND.frx":23EFC
         Left            =   -73320
         List            =   "frmCourseHND.frx":23F03
         TabIndex        =   98
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List17 
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
         ItemData        =   "frmCourseHND.frx":23F15
         Left            =   -68280
         List            =   "frmCourseHND.frx":23F1C
         TabIndex        =   97
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List16 
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
         ItemData        =   "frmCourseHND.frx":23F23
         Left            =   -67920
         List            =   "frmCourseHND.frx":23F2A
         TabIndex        =   96
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List15 
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
         ItemData        =   "frmCourseHND.frx":23F31
         Left            =   -67560
         List            =   "frmCourseHND.frx":23F38
         TabIndex        =   95
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHND.frx":23F40
         Left            =   -66960
         List            =   "frmCourseHND.frx":23F47
         TabIndex        =   94
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit4 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   93
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack4 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   92
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack3 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   87
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit3 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   86
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack2 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -68760
         TabIndex        =   85
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit2 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70200
         TabIndex        =   84
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack1 
         Caption         =   "&Back"
         Height          =   495
         Left            =   5880
         TabIndex        =   83
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit1 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   4440
         TabIndex        =   82
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   -69120
         TabIndex        =   81
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   -70560
         TabIndex        =   80
         Top             =   4920
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
         ItemData        =   "frmCourseHND.frx":23F51
         Left            =   -66120
         List            =   "frmCourseHND.frx":23F58
         TabIndex        =   73
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List12 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23F67
         Left            =   -74760
         List            =   "frmCourseHND.frx":23F80
         TabIndex        =   72
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List11 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":23FB6
         Left            =   -73320
         List            =   "frmCourseHND.frx":23FCF
         TabIndex        =   71
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List10 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2408F
         Left            =   -68280
         List            =   "frmCourseHND.frx":240A8
         TabIndex        =   70
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List9 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":240C1
         Left            =   -67920
         List            =   "frmCourseHND.frx":240DA
         TabIndex        =   69
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List8 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":240F4
         Left            =   -67560
         List            =   "frmCourseHND.frx":24110
         TabIndex        =   68
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List7 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2412C
         Left            =   -66960
         List            =   "frmCourseHND.frx":24145
         TabIndex        =   67
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List6 
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
         ItemData        =   "frmCourseHND.frx":24164
         Left            =   -74760
         List            =   "frmCourseHND.frx":2416B
         TabIndex        =   66
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List5 
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
         ItemData        =   "frmCourseHND.frx":24178
         Left            =   -73320
         List            =   "frmCourseHND.frx":2417F
         TabIndex        =   65
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List4 
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
         ItemData        =   "frmCourseHND.frx":24191
         Left            =   -68280
         List            =   "frmCourseHND.frx":24198
         TabIndex        =   64
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List3 
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
         ItemData        =   "frmCourseHND.frx":2419F
         Left            =   -67920
         List            =   "frmCourseHND.frx":241A6
         TabIndex        =   63
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List2 
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
         ItemData        =   "frmCourseHND.frx":241AD
         Left            =   -67560
         List            =   "frmCourseHND.frx":241B4
         TabIndex        =   62
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List1 
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
         ItemData        =   "frmCourseHND.frx":241BC
         Left            =   -66960
         List            =   "frmCourseHND.frx":241C3
         TabIndex        =   61
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List131 
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
         ItemData        =   "frmCourseHND.frx":241CD
         Left            =   -66960
         List            =   "frmCourseHND.frx":241D4
         TabIndex        =   54
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List132 
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
         ItemData        =   "frmCourseHND.frx":241DE
         Left            =   -67560
         List            =   "frmCourseHND.frx":241E5
         TabIndex        =   53
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List133 
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
         ItemData        =   "frmCourseHND.frx":241ED
         Left            =   -67920
         List            =   "frmCourseHND.frx":241F4
         TabIndex        =   52
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List134 
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
         ItemData        =   "frmCourseHND.frx":241FB
         Left            =   -68280
         List            =   "frmCourseHND.frx":24202
         TabIndex        =   51
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List135 
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
         ItemData        =   "frmCourseHND.frx":24209
         Left            =   -73320
         List            =   "frmCourseHND.frx":24210
         TabIndex        =   50
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List136 
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
         ItemData        =   "frmCourseHND.frx":24222
         Left            =   -74760
         List            =   "frmCourseHND.frx":24229
         TabIndex        =   49
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List137 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":24236
         Left            =   -66960
         List            =   "frmCourseHND.frx":2424C
         TabIndex        =   48
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List138 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":24269
         Left            =   -67560
         List            =   "frmCourseHND.frx":24282
         TabIndex        =   47
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List139 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":2429C
         Left            =   -67920
         List            =   "frmCourseHND.frx":242B2
         TabIndex        =   46
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List140 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":242C8
         Left            =   -68280
         List            =   "frmCourseHND.frx":242DE
         TabIndex        =   45
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List141 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":242F4
         Left            =   -73320
         List            =   "frmCourseHND.frx":2430A
         TabIndex        =   44
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List142 
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
         Height          =   2910
         ItemData        =   "frmCourseHND.frx":243AE
         Left            =   -74760
         List            =   "frmCourseHND.frx":243C4
         TabIndex        =   43
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List143 
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
         ItemData        =   "frmCourseHND.frx":243F8
         Left            =   -66120
         List            =   "frmCourseHND.frx":243FF
         TabIndex        =   42
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List118 
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
         ItemData        =   "frmCourseHND.frx":2440E
         Left            =   8040
         List            =   "frmCourseHND.frx":24415
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List119 
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
         ItemData        =   "frmCourseHND.frx":2441F
         Left            =   7440
         List            =   "frmCourseHND.frx":24426
         TabIndex        =   33
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List120 
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
         ItemData        =   "frmCourseHND.frx":2442E
         Left            =   7080
         List            =   "frmCourseHND.frx":24435
         TabIndex        =   32
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List121 
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
         ItemData        =   "frmCourseHND.frx":2443C
         Left            =   6720
         List            =   "frmCourseHND.frx":24443
         TabIndex        =   31
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List122 
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
         ItemData        =   "frmCourseHND.frx":2444A
         Left            =   1680
         List            =   "frmCourseHND.frx":24451
         TabIndex        =   30
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List123 
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
         ItemData        =   "frmCourseHND.frx":24463
         Left            =   240
         List            =   "frmCourseHND.frx":2446A
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List124 
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
         ItemData        =   "frmCourseHND.frx":24477
         Left            =   8040
         List            =   "frmCourseHND.frx":24490
         TabIndex        =   28
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List125 
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
         ItemData        =   "frmCourseHND.frx":244B2
         Left            =   7440
         List            =   "frmCourseHND.frx":244CE
         TabIndex        =   27
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List126 
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
         ItemData        =   "frmCourseHND.frx":244EB
         Left            =   7080
         List            =   "frmCourseHND.frx":24504
         TabIndex        =   26
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List127 
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
         ItemData        =   "frmCourseHND.frx":2451D
         Left            =   6720
         List            =   "frmCourseHND.frx":24536
         TabIndex        =   25
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List128 
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
         ItemData        =   "frmCourseHND.frx":2454F
         Left            =   1680
         List            =   "frmCourseHND.frx":24568
         TabIndex        =   24
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List129 
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
         ItemData        =   "frmCourseHND.frx":2461C
         Left            =   240
         List            =   "frmCourseHND.frx":24635
         TabIndex        =   23
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List130 
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
         ItemData        =   "frmCourseHND.frx":24671
         Left            =   8880
         List            =   "frmCourseHND.frx":24678
         TabIndex        =   22
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox List105 
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
         ItemData        =   "frmCourseHND.frx":24687
         Left            =   -66960
         List            =   "frmCourseHND.frx":2468E
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List106 
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
         ItemData        =   "frmCourseHND.frx":24698
         Left            =   -67560
         List            =   "frmCourseHND.frx":2469F
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List107 
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
         ItemData        =   "frmCourseHND.frx":246A7
         Left            =   -67920
         List            =   "frmCourseHND.frx":246AE
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List108 
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
         ItemData        =   "frmCourseHND.frx":246B5
         Left            =   -68280
         List            =   "frmCourseHND.frx":246BC
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List109 
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
         ItemData        =   "frmCourseHND.frx":246C3
         Left            =   -73320
         List            =   "frmCourseHND.frx":246CA
         TabIndex        =   10
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List110 
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
         ItemData        =   "frmCourseHND.frx":246DC
         Left            =   -74760
         List            =   "frmCourseHND.frx":246E3
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List111 
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
         ItemData        =   "frmCourseHND.frx":246F0
         Left            =   -66960
         List            =   "frmCourseHND.frx":24709
         TabIndex        =   8
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List112 
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
         ItemData        =   "frmCourseHND.frx":2472C
         Left            =   -67560
         List            =   "frmCourseHND.frx":24748
         TabIndex        =   7
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List113 
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
         ItemData        =   "frmCourseHND.frx":24765
         Left            =   -67920
         List            =   "frmCourseHND.frx":2477E
         TabIndex        =   6
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List114 
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
         ItemData        =   "frmCourseHND.frx":24797
         Left            =   -68280
         List            =   "frmCourseHND.frx":247B0
         TabIndex        =   5
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List115 
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
         ItemData        =   "frmCourseHND.frx":247C9
         Left            =   -73320
         List            =   "frmCourseHND.frx":247E2
         TabIndex        =   4
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List116 
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
         ItemData        =   "frmCourseHND.frx":2488A
         Left            =   -74760
         List            =   "frmCourseHND.frx":248A3
         TabIndex        =   3
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox List117 
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
         ItemData        =   "frmCourseHND.frx":248DF
         Left            =   -66120
         List            =   "frmCourseHND.frx":248E6
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   -65280
         Top             =   4440
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
         RecordSource    =   "tblHND1FirstS"
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
      Begin MSAdodcLib.Adodc Adodc2 
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
         RecordSource    =   "tblHND1SecondS"
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
         RecordSource    =   "tblHND2FirstS"
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
         RecordSource    =   "tblHND2SecondS"
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
         Left            =   -65280
         Top             =   4440
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
         RecordSource    =   "tblHND1EFirstS"
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
         RecordSource    =   "tblHND1ESecondS"
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
         RecordSource    =   "tblHND2EFirstS"
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
         RecordSource    =   "tblHND2ESecondS"
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
      Begin VB.Label Label53 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM422"
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
         TabIndex        =   181
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM423"
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
         TabIndex        =   180
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM424"
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
         TabIndex        =   179
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM426"
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
         TabIndex        =   178
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFFF&
         DataField       =   "EED413"
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
         TabIndex        =   177
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM429"
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
         TabIndex        =   176
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "HND I (EVENING) SECOND SEMESTER"
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
         Left            =   -72000
         TabIndex        =   175
         Top             =   480
         Width           =   6000
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA411"
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
         Left            =   -66120
         TabIndex        =   159
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM416"
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
         Left            =   -66120
         TabIndex        =   158
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM415"
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
         Left            =   -66120
         TabIndex        =   157
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM414"
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
         Left            =   -66120
         TabIndex        =   156
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM413"
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
         Left            =   -66120
         TabIndex        =   155
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM412"
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
         Left            =   -66120
         TabIndex        =   154
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label G 
         AutoSize        =   -1  'True
         Caption         =   "HND II (EVENING) FIRST SEMESTER"
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
         Left            =   -71880
         TabIndex        =   153
         Top             =   480
         Width           =   5805
      End
      Begin VB.Label F 
         AutoSize        =   -1  'True
         Caption         =   "HND I (EVENING) SECOND SEMESTER"
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
         Left            =   -71880
         TabIndex        =   137
         Top             =   480
         Width           =   6000
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM320"
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
         TabIndex        =   136
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA321"
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
         TabIndex        =   135
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM326"
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
         TabIndex        =   134
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM325"
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
         TabIndex        =   133
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM323"
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
         TabIndex        =   132
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM322"
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
         TabIndex        =   131
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM321"
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
         TabIndex        =   130
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM315"
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
         TabIndex        =   114
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sta314"
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
         TabIndex        =   113
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sta311"
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
         TabIndex        =   112
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com314"
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
         TabIndex        =   111
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com313"
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
         TabIndex        =   110
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com312"
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
         TabIndex        =   109
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com311"
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
         TabIndex        =   108
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label E 
         AutoSize        =   -1  'True
         Caption         =   "HND I (EVENING) FIRST SEMESTER"
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
         Left            =   -71880
         TabIndex        =   107
         Top             =   480
         Width           =   5700
      End
      Begin VB.Label D 
         AutoSize        =   -1  'True
         Caption         =   "HND I SECOND SEMESTER"
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
         Left            =   -71400
         TabIndex        =   91
         Top             =   480
         Width           =   4305
      End
      Begin VB.Label B 
         AutoSize        =   -1  'True
         Caption         =   "HND I SECOND SEMESTER"
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
         Left            =   3480
         TabIndex        =   90
         Top             =   480
         Width           =   4305
      End
      Begin VB.Label C 
         AutoSize        =   -1  'True
         Caption         =   "HND II FIRST SEMESTER"
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
         Left            =   -71520
         TabIndex        =   89
         Top             =   480
         Width           =   4110
      End
      Begin VB.Label A 
         AutoSize        =   -1  'True
         Caption         =   "HND I FIRST SEMESTER"
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
         Left            =   -71280
         TabIndex        =   88
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM429"
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
         TabIndex        =   79
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         DataField       =   "EED413"
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
         TabIndex        =   78
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM426"
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
         TabIndex        =   77
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM424"
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
         TabIndex        =   76
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM423"
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
         TabIndex        =   75
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM422"
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
         TabIndex        =   74
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM412"
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
         TabIndex        =   60
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM413"
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
         TabIndex        =   59
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM414"
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
         TabIndex        =   58
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM415"
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
         TabIndex        =   57
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM416"
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
         TabIndex        =   56
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA411"
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
         TabIndex        =   55
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM321"
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
         Left            =   8880
         TabIndex        =   41
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM322"
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
         Left            =   8880
         TabIndex        =   40
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM323"
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
         Left            =   8880
         TabIndex        =   39
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM325"
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
         Left            =   8880
         TabIndex        =   38
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COM326"
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
         Left            =   8880
         TabIndex        =   37
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         DataField       =   "STA321"
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
         Left            =   8880
         TabIndex        =   36
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM320"
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
         Left            =   8880
         TabIndex        =   35
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com311"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com312"
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
         TabIndex        =   20
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com313"
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
         TabIndex        =   19
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "com314"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sta311"
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
         TabIndex        =   17
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "sta314"
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
         TabIndex        =   16
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OTM315"
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
         TabIndex        =   15
         Top             =   3480
         Width           =   2895
      End
   End
   Begin VB.Label logo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTAL COURSES for HND"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   2880
      Width           =   6240
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmCourseHND.frx":248F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmCourseHND"
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
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 0
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
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 1
End Sub

Private Sub cmdEdit2_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 2
End Sub

Private Sub cmdEdit3_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 3
End Sub

Private Sub cmdEdit4_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 4
End Sub

Private Sub cmdEdit5_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 5
End Sub

Private Sub cmdEdit6_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 6
End Sub

Private Sub cmdEdit7_Click()
Me.Hide
frmCourseHNDEdit.Show
frmCourseHNDEdit.SSTab1.Tab = 7
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
