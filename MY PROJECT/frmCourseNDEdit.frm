VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCourseNDEdit 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Courses and their Lecturers Editor"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12345
   Icon            =   "frmCourseNDEdit.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   77
      Top             =   3600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "ND1(M)1S"
      TabPicture(0)   =   "frmCourseNDEdit.frx":234CD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "A"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Adodc1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "List11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "List9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "List8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "List7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "List6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "List5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "List4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "List3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "List2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "List1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "List13"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdOK"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "ND1(M)2S"
      TabPicture(1)   =   "frmCourseNDEdit.frx":234E9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "B"
      Tab(1).Control(1)=   "Adodc2"
      Tab(1).Control(2)=   "Text19"
      Tab(1).Control(3)=   "Text18"
      Tab(1).Control(4)=   "Text17"
      Tab(1).Control(5)=   "Text16"
      Tab(1).Control(6)=   "Text15"
      Tab(1).Control(7)=   "Text14"
      Tab(1).Control(8)=   "Text13"
      Tab(1).Control(9)=   "Text12"
      Tab(1).Control(10)=   "Text11"
      Tab(1).Control(11)=   "Text10"
      Tab(1).Control(12)=   "List26"
      Tab(1).Control(13)=   "List25"
      Tab(1).Control(14)=   "List24"
      Tab(1).Control(15)=   "List23"
      Tab(1).Control(16)=   "List22"
      Tab(1).Control(17)=   "List21"
      Tab(1).Control(18)=   "List20"
      Tab(1).Control(19)=   "List19"
      Tab(1).Control(20)=   "List18"
      Tab(1).Control(21)=   "List17"
      Tab(1).Control(22)=   "List16"
      Tab(1).Control(23)=   "List15"
      Tab(1).Control(24)=   "List14"
      Tab(1).Control(25)=   "cmdOK1"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "ND1(E)1S"
      TabPicture(2)   =   "frmCourseNDEdit.frx":23505
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label29"
      Tab(2).Control(1)=   "Adodc3"
      Tab(2).Control(2)=   "Text28"
      Tab(2).Control(3)=   "Text27"
      Tab(2).Control(4)=   "Text26"
      Tab(2).Control(5)=   "Text25"
      Tab(2).Control(6)=   "Text24"
      Tab(2).Control(7)=   "Text23"
      Tab(2).Control(8)=   "Text22"
      Tab(2).Control(9)=   "Text21"
      Tab(2).Control(10)=   "Text20"
      Tab(2).Control(11)=   "List39"
      Tab(2).Control(12)=   "List38"
      Tab(2).Control(13)=   "List37"
      Tab(2).Control(14)=   "List36"
      Tab(2).Control(15)=   "List35"
      Tab(2).Control(16)=   "List34"
      Tab(2).Control(17)=   "List33"
      Tab(2).Control(18)=   "List32"
      Tab(2).Control(19)=   "List31"
      Tab(2).Control(20)=   "List30"
      Tab(2).Control(21)=   "List29"
      Tab(2).Control(22)=   "List28"
      Tab(2).Control(23)=   "List27"
      Tab(2).Control(24)=   "cmdOK2"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "ND1(E)2S"
      TabPicture(3)   =   "frmCourseNDEdit.frx":23521
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label70"
      Tab(3).Control(1)=   "Adodc4"
      Tab(3).Control(2)=   "Text38"
      Tab(3).Control(3)=   "Text37"
      Tab(3).Control(4)=   "Text36"
      Tab(3).Control(5)=   "Text35"
      Tab(3).Control(6)=   "Text34"
      Tab(3).Control(7)=   "Text33"
      Tab(3).Control(8)=   "Text32"
      Tab(3).Control(9)=   "Text31"
      Tab(3).Control(10)=   "Text30"
      Tab(3).Control(11)=   "Text29"
      Tab(3).Control(12)=   "List52"
      Tab(3).Control(13)=   "List51"
      Tab(3).Control(14)=   "List50"
      Tab(3).Control(15)=   "List49"
      Tab(3).Control(16)=   "List48"
      Tab(3).Control(17)=   "List47"
      Tab(3).Control(18)=   "List46"
      Tab(3).Control(19)=   "List45"
      Tab(3).Control(20)=   "List44"
      Tab(3).Control(21)=   "List43"
      Tab(3).Control(22)=   "List42"
      Tab(3).Control(23)=   "List41"
      Tab(3).Control(24)=   "List40"
      Tab(3).Control(25)=   "cmdOK3"
      Tab(3).ControlCount=   26
      TabCaption(4)   =   "ND2(M)1S"
      TabPicture(4)   =   "frmCourseNDEdit.frx":2353D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label71"
      Tab(4).Control(1)=   "Adodc5"
      Tab(4).Control(2)=   "Text45"
      Tab(4).Control(3)=   "Text44"
      Tab(4).Control(4)=   "Text43"
      Tab(4).Control(5)=   "Text42"
      Tab(4).Control(6)=   "Text41"
      Tab(4).Control(7)=   "Text40"
      Tab(4).Control(8)=   "Text39"
      Tab(4).Control(9)=   "List65"
      Tab(4).Control(10)=   "List64"
      Tab(4).Control(11)=   "List63"
      Tab(4).Control(12)=   "List62"
      Tab(4).Control(13)=   "List61"
      Tab(4).Control(14)=   "List60"
      Tab(4).Control(15)=   "List59"
      Tab(4).Control(16)=   "List58"
      Tab(4).Control(17)=   "List57"
      Tab(4).Control(18)=   "List56"
      Tab(4).Control(19)=   "List55"
      Tab(4).Control(20)=   "List54"
      Tab(4).Control(21)=   "List53"
      Tab(4).Control(22)=   "cmdOK4"
      Tab(4).ControlCount=   23
      TabCaption(5)   =   "ND2(M)2S"
      TabPicture(5)   =   "frmCourseNDEdit.frx":23559
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label72"
      Tab(5).Control(1)=   "Adodc6"
      Tab(5).Control(2)=   "Text53"
      Tab(5).Control(3)=   "Text52"
      Tab(5).Control(4)=   "Text51"
      Tab(5).Control(5)=   "Text50"
      Tab(5).Control(6)=   "Text49"
      Tab(5).Control(7)=   "Text48"
      Tab(5).Control(8)=   "Text47"
      Tab(5).Control(9)=   "Text46"
      Tab(5).Control(10)=   "List78"
      Tab(5).Control(11)=   "List77"
      Tab(5).Control(12)=   "List76"
      Tab(5).Control(13)=   "List75"
      Tab(5).Control(14)=   "List74"
      Tab(5).Control(15)=   "List73"
      Tab(5).Control(16)=   "List72"
      Tab(5).Control(17)=   "List71"
      Tab(5).Control(18)=   "List70"
      Tab(5).Control(19)=   "List69"
      Tab(5).Control(20)=   "List68"
      Tab(5).Control(21)=   "List67"
      Tab(5).Control(22)=   "List66"
      Tab(5).Control(23)=   "cmdOK5"
      Tab(5).ControlCount=   24
      TabCaption(6)   =   "ND2(E)1S"
      TabPicture(6)   =   "frmCourseNDEdit.frx":23575
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label73"
      Tab(6).Control(1)=   "Adodc7"
      Tab(6).Control(2)=   "Text60"
      Tab(6).Control(3)=   "Text59"
      Tab(6).Control(4)=   "Text58"
      Tab(6).Control(5)=   "Text57"
      Tab(6).Control(6)=   "Text56"
      Tab(6).Control(7)=   "Text55"
      Tab(6).Control(8)=   "Text54"
      Tab(6).Control(9)=   "List91"
      Tab(6).Control(10)=   "List90"
      Tab(6).Control(11)=   "List89"
      Tab(6).Control(12)=   "List88"
      Tab(6).Control(13)=   "List87"
      Tab(6).Control(14)=   "List86"
      Tab(6).Control(15)=   "List85"
      Tab(6).Control(16)=   "List84"
      Tab(6).Control(17)=   "List83"
      Tab(6).Control(18)=   "List82"
      Tab(6).Control(19)=   "List81"
      Tab(6).Control(20)=   "List80"
      Tab(6).Control(21)=   "List79"
      Tab(6).Control(22)=   "cmdOK6"
      Tab(6).ControlCount=   23
      TabCaption(7)   =   "ND2(E)2S"
      TabPicture(7)   =   "frmCourseNDEdit.frx":23591
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label74"
      Tab(7).Control(1)=   "Adodc8"
      Tab(7).Control(2)=   "Text68"
      Tab(7).Control(3)=   "Text67"
      Tab(7).Control(4)=   "Text66"
      Tab(7).Control(5)=   "Text65"
      Tab(7).Control(6)=   "Text64"
      Tab(7).Control(7)=   "Text63"
      Tab(7).Control(8)=   "Text62"
      Tab(7).Control(9)=   "Text61"
      Tab(7).Control(10)=   "List104"
      Tab(7).Control(11)=   "List103"
      Tab(7).Control(12)=   "List102"
      Tab(7).Control(13)=   "List101"
      Tab(7).Control(14)=   "List100"
      Tab(7).Control(15)=   "List99"
      Tab(7).Control(16)=   "List98"
      Tab(7).Control(17)=   "List97"
      Tab(7).Control(18)=   "List96"
      Tab(7).Control(19)=   "List95"
      Tab(7).Control(20)=   "List94"
      Tab(7).Control(21)=   "List93"
      Tab(7).Control(22)=   "List92"
      Tab(7).Control(23)=   "cmdOK7"
      Tab(7).ControlCount=   24
      Begin VB.CommandButton cmdOK7 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   76
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK6 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   67
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK5 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   59
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   50
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK3 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   42
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK2 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   31
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK1 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   21
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   5280
         TabIndex        =   10
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
         ItemData        =   "frmCourseNDEdit.frx":235AD
         Left            =   8880
         List            =   "frmCourseNDEdit.frx":235B4
         TabIndex        =   181
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
         ItemData        =   "frmCourseNDEdit.frx":235C3
         Left            =   240
         List            =   "frmCourseNDEdit.frx":235E2
         TabIndex        =   180
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
         ItemData        =   "frmCourseNDEdit.frx":23630
         Left            =   1680
         List            =   "frmCourseNDEdit.frx":2364F
         TabIndex        =   179
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
         ItemData        =   "frmCourseNDEdit.frx":2374E
         Left            =   6720
         List            =   "frmCourseNDEdit.frx":2376D
         TabIndex        =   178
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
         ItemData        =   "frmCourseNDEdit.frx":2378C
         Left            =   7080
         List            =   "frmCourseNDEdit.frx":237AB
         TabIndex        =   177
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
         ItemData        =   "frmCourseNDEdit.frx":237CA
         Left            =   7440
         List            =   "frmCourseNDEdit.frx":237EC
         TabIndex        =   176
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
         ItemData        =   "frmCourseNDEdit.frx":2380F
         Left            =   8040
         List            =   "frmCourseNDEdit.frx":2382E
         TabIndex        =   175
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
         ItemData        =   "frmCourseNDEdit.frx":2385C
         Left            =   240
         List            =   "frmCourseNDEdit.frx":23863
         TabIndex        =   174
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
         ItemData        =   "frmCourseNDEdit.frx":23870
         Left            =   1680
         List            =   "frmCourseNDEdit.frx":23877
         TabIndex        =   173
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
         ItemData        =   "frmCourseNDEdit.frx":23889
         Left            =   6720
         List            =   "frmCourseNDEdit.frx":23890
         TabIndex        =   172
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
         ItemData        =   "frmCourseNDEdit.frx":23897
         Left            =   7080
         List            =   "frmCourseNDEdit.frx":2389E
         TabIndex        =   171
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
         ItemData        =   "frmCourseNDEdit.frx":238A5
         Left            =   7440
         List            =   "frmCourseNDEdit.frx":238AC
         TabIndex        =   170
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
         ItemData        =   "frmCourseNDEdit.frx":238B4
         Left            =   8040
         List            =   "frmCourseNDEdit.frx":238BB
         TabIndex        =   169
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
         ItemData        =   "frmCourseNDEdit.frx":238C5
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":238CC
         TabIndex        =   168
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
         ItemData        =   "frmCourseNDEdit.frx":238DB
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":238FD
         TabIndex        =   167
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
         ItemData        =   "frmCourseNDEdit.frx":23951
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":23973
         TabIndex        =   166
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
         ItemData        =   "frmCourseNDEdit.frx":23AA3
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":23AC5
         TabIndex        =   165
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
         ItemData        =   "frmCourseNDEdit.frx":23AE7
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":23B09
         TabIndex        =   164
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
         ItemData        =   "frmCourseNDEdit.frx":23B2B
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":23B50
         TabIndex        =   163
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
         ItemData        =   "frmCourseNDEdit.frx":23B76
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23B98
         TabIndex        =   162
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
         ItemData        =   "frmCourseNDEdit.frx":23BC8
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":23BCF
         TabIndex        =   161
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
         ItemData        =   "frmCourseNDEdit.frx":23BDC
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":23BE3
         TabIndex        =   160
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
         ItemData        =   "frmCourseNDEdit.frx":23BF5
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":23BFC
         TabIndex        =   159
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
         ItemData        =   "frmCourseNDEdit.frx":23C03
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":23C0A
         TabIndex        =   158
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
         ItemData        =   "frmCourseNDEdit.frx":23C11
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":23C18
         TabIndex        =   157
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
         ItemData        =   "frmCourseNDEdit.frx":23C20
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23C27
         TabIndex        =   156
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
         ItemData        =   "frmCourseNDEdit.frx":23C31
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23C38
         TabIndex        =   155
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
         ItemData        =   "frmCourseNDEdit.frx":23C42
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":23C49
         TabIndex        =   154
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
         ItemData        =   "frmCourseNDEdit.frx":23C51
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":23C58
         TabIndex        =   153
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
         ItemData        =   "frmCourseNDEdit.frx":23C5F
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":23C66
         TabIndex        =   152
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
         ItemData        =   "frmCourseNDEdit.frx":23C6D
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":23C74
         TabIndex        =   151
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
         ItemData        =   "frmCourseNDEdit.frx":23C86
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":23C8D
         TabIndex        =   150
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
         ItemData        =   "frmCourseNDEdit.frx":23C9A
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23CB9
         TabIndex        =   149
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
         ItemData        =   "frmCourseNDEdit.frx":23CE7
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":23D09
         TabIndex        =   148
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
         ItemData        =   "frmCourseNDEdit.frx":23D2C
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":23D4B
         TabIndex        =   147
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
         ItemData        =   "frmCourseNDEdit.frx":23D6A
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":23D89
         TabIndex        =   146
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
         ItemData        =   "frmCourseNDEdit.frx":23DA8
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":23DC7
         TabIndex        =   145
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
         ItemData        =   "frmCourseNDEdit.frx":23EC6
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":23EE5
         TabIndex        =   144
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
         ItemData        =   "frmCourseNDEdit.frx":23F33
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":23F3A
         TabIndex        =   143
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
         ItemData        =   "frmCourseNDEdit.frx":23F49
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23F50
         TabIndex        =   142
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
         ItemData        =   "frmCourseNDEdit.frx":23F5A
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":23F61
         TabIndex        =   141
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
         ItemData        =   "frmCourseNDEdit.frx":23F69
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":23F70
         TabIndex        =   140
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
         ItemData        =   "frmCourseNDEdit.frx":23F77
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":23F7E
         TabIndex        =   139
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
         ItemData        =   "frmCourseNDEdit.frx":23F85
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":23F8C
         TabIndex        =   138
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
         ItemData        =   "frmCourseNDEdit.frx":23F9E
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":23FA5
         TabIndex        =   137
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
         ItemData        =   "frmCourseNDEdit.frx":23FB2
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":23FD4
         TabIndex        =   136
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
         ItemData        =   "frmCourseNDEdit.frx":24004
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24029
         TabIndex        =   135
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
         ItemData        =   "frmCourseNDEdit.frx":2404F
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":24071
         TabIndex        =   134
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
         ItemData        =   "frmCourseNDEdit.frx":24093
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":240B5
         TabIndex        =   133
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
         ItemData        =   "frmCourseNDEdit.frx":240D7
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":240F9
         TabIndex        =   132
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
         ItemData        =   "frmCourseNDEdit.frx":24229
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":2424B
         TabIndex        =   131
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
         ItemData        =   "frmCourseNDEdit.frx":2429F
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":242A6
         TabIndex        =   130
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
         ItemData        =   "frmCourseNDEdit.frx":242B5
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":242BC
         TabIndex        =   129
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
         ItemData        =   "frmCourseNDEdit.frx":242C6
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":242CD
         TabIndex        =   128
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
         ItemData        =   "frmCourseNDEdit.frx":242D5
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":242DC
         TabIndex        =   127
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
         ItemData        =   "frmCourseNDEdit.frx":242E3
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":242EA
         TabIndex        =   126
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
         ItemData        =   "frmCourseNDEdit.frx":242F1
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":242F8
         TabIndex        =   125
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
         ItemData        =   "frmCourseNDEdit.frx":2430A
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":24311
         TabIndex        =   124
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
         ItemData        =   "frmCourseNDEdit.frx":2431E
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":24337
         TabIndex        =   123
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
         ItemData        =   "frmCourseNDEdit.frx":24358
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24374
         TabIndex        =   122
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
         ItemData        =   "frmCourseNDEdit.frx":24391
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":243AA
         TabIndex        =   121
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
         ItemData        =   "frmCourseNDEdit.frx":243C3
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":243DC
         TabIndex        =   120
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
         ItemData        =   "frmCourseNDEdit.frx":243F5
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":2440E
         TabIndex        =   119
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
         ItemData        =   "frmCourseNDEdit.frx":244D3
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":244EC
         TabIndex        =   118
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
         ItemData        =   "frmCourseNDEdit.frx":24528
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":2452F
         TabIndex        =   117
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
         ItemData        =   "frmCourseNDEdit.frx":2453E
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":24545
         TabIndex        =   116
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
         ItemData        =   "frmCourseNDEdit.frx":2454F
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24556
         TabIndex        =   115
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
         ItemData        =   "frmCourseNDEdit.frx":2455E
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":24565
         TabIndex        =   114
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
         ItemData        =   "frmCourseNDEdit.frx":2456C
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":24573
         TabIndex        =   113
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
         ItemData        =   "frmCourseNDEdit.frx":2457A
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":24581
         TabIndex        =   112
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
         ItemData        =   "frmCourseNDEdit.frx":24593
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":2459A
         TabIndex        =   111
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
         ItemData        =   "frmCourseNDEdit.frx":245A7
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":245C3
         TabIndex        =   110
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
         ItemData        =   "frmCourseNDEdit.frx":245E8
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24607
         TabIndex        =   109
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
         ItemData        =   "frmCourseNDEdit.frx":24627
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":24643
         TabIndex        =   108
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
         ItemData        =   "frmCourseNDEdit.frx":2465F
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":2467B
         TabIndex        =   107
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
         ItemData        =   "frmCourseNDEdit.frx":24697
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":246B3
         TabIndex        =   106
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
         ItemData        =   "frmCourseNDEdit.frx":24780
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":2479C
         TabIndex        =   105
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
         ItemData        =   "frmCourseNDEdit.frx":247E0
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":247E7
         TabIndex        =   104
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
         ItemData        =   "frmCourseNDEdit.frx":247F6
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":247FD
         TabIndex        =   103
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
         ItemData        =   "frmCourseNDEdit.frx":24807
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":2480E
         TabIndex        =   102
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
         ItemData        =   "frmCourseNDEdit.frx":24816
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":2481D
         TabIndex        =   101
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
         ItemData        =   "frmCourseNDEdit.frx":24824
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":2482B
         TabIndex        =   100
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
         ItemData        =   "frmCourseNDEdit.frx":24832
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":24839
         TabIndex        =   99
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
         ItemData        =   "frmCourseNDEdit.frx":2484B
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":24852
         TabIndex        =   98
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
         ItemData        =   "frmCourseNDEdit.frx":2485F
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":24878
         TabIndex        =   97
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
         ItemData        =   "frmCourseNDEdit.frx":24899
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":248B5
         TabIndex        =   96
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
         ItemData        =   "frmCourseNDEdit.frx":248D2
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":248EB
         TabIndex        =   95
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
         ItemData        =   "frmCourseNDEdit.frx":24904
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":2491D
         TabIndex        =   94
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
         ItemData        =   "frmCourseNDEdit.frx":24936
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":2494F
         TabIndex        =   93
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
         ItemData        =   "frmCourseNDEdit.frx":24A14
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":24A2D
         TabIndex        =   92
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
         ItemData        =   "frmCourseNDEdit.frx":24A69
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":24A70
         TabIndex        =   91
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
         ItemData        =   "frmCourseNDEdit.frx":24A7F
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":24A86
         TabIndex        =   90
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
         ItemData        =   "frmCourseNDEdit.frx":24A90
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24A97
         TabIndex        =   89
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
         ItemData        =   "frmCourseNDEdit.frx":24A9F
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":24AA6
         TabIndex        =   88
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
         ItemData        =   "frmCourseNDEdit.frx":24AAD
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":24AB4
         TabIndex        =   87
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
         ItemData        =   "frmCourseNDEdit.frx":24ABB
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":24AC2
         TabIndex        =   86
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
         ItemData        =   "frmCourseNDEdit.frx":24AD4
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":24ADB
         TabIndex        =   85
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
         ItemData        =   "frmCourseNDEdit.frx":24AE8
         Left            =   -66960
         List            =   "frmCourseNDEdit.frx":24B04
         TabIndex        =   84
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
         ItemData        =   "frmCourseNDEdit.frx":24B29
         Left            =   -67560
         List            =   "frmCourseNDEdit.frx":24B48
         TabIndex        =   83
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
         ItemData        =   "frmCourseNDEdit.frx":24B68
         Left            =   -67920
         List            =   "frmCourseNDEdit.frx":24B84
         TabIndex        =   82
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
         ItemData        =   "frmCourseNDEdit.frx":24BA0
         Left            =   -68280
         List            =   "frmCourseNDEdit.frx":24BBC
         TabIndex        =   81
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
         ItemData        =   "frmCourseNDEdit.frx":24BD8
         Left            =   -73320
         List            =   "frmCourseNDEdit.frx":24BF4
         TabIndex        =   80
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
         ItemData        =   "frmCourseNDEdit.frx":24CC1
         Left            =   -74760
         List            =   "frmCourseNDEdit.frx":24CDD
         TabIndex        =   79
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
         ItemData        =   "frmCourseNDEdit.frx":24D21
         Left            =   -66120
         List            =   "frmCourseNDEdit.frx":24D28
         TabIndex        =   78
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "COM101"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         DataField       =   "COM121"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text11 
         DataField       =   "COM122"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   12
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text12 
         DataField       =   "COM123"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   13
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         DataField       =   "COM124"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   14
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         DataField       =   "COM125"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   15
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         DataField       =   "COM126"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   16
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text16 
         DataField       =   "GNS102"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   17
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text17 
         DataField       =   "GNS128"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   18
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text18 
         DataField       =   "EED126"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   19
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox Text19 
         DataField       =   "URP120"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   20
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox Text20 
         DataField       =   "COM101"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   22
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text21 
         DataField       =   "COM112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   23
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text22 
         DataField       =   "COM113"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   24
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text23 
         DataField       =   "STA111"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   25
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text24 
         DataField       =   "STA112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   26
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text25 
         DataField       =   "MATH111"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   27
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text26 
         DataField       =   "MATH112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   28
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text27 
         DataField       =   "OTM112"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   29
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text28 
         DataField       =   "GNS127"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   30
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox Text29 
         DataField       =   "COM121"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   32
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text30 
         DataField       =   "COM122"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   33
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text31 
         DataField       =   "COM123"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   34
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text32 
         DataField       =   "COM124"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   35
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text33 
         DataField       =   "COM125"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   36
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text34 
         DataField       =   "COM126"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   37
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text35 
         DataField       =   "GNS102"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   38
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text36 
         DataField       =   "GNS128"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   39
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text37 
         DataField       =   "EED126"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   40
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox Text38 
         DataField       =   "URP120"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   41
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox Text39 
         DataField       =   "COM211"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   43
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text40 
         DataField       =   "COM212"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   44
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text41 
         DataField       =   "COM213"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   45
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text42 
         DataField       =   "COM214"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   46
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text43 
         DataField       =   "COM215"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   47
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text44 
         DataField       =   "COM216"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   48
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text45 
         DataField       =   "GNS201"
         DataSource      =   "Adodc5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   49
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text46 
         DataField       =   "COM221"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   51
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text47 
         DataField       =   "COM222"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   52
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text48 
         DataField       =   "COM223"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   53
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text49 
         DataField       =   "COM224"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   54
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text50 
         DataField       =   "COM225"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   55
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text51 
         DataField       =   "COM226"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   56
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text52 
         DataField       =   "STA226"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   57
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text53 
         DataField       =   "COM229"
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   58
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text54 
         DataField       =   "COM211"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   60
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text55 
         DataField       =   "COM212"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   61
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text56 
         DataField       =   "COM213"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   62
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text57 
         DataField       =   "COM214"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   63
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text58 
         DataField       =   "COM215"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   64
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text59 
         DataField       =   "COM216"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   65
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text60 
         DataField       =   "GNS201"
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   66
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text61 
         DataField       =   "COM221"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   68
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text62 
         DataField       =   "COM222"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   69
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text63 
         DataField       =   "COM223"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   70
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text64 
         DataField       =   "COM224"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   71
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text65 
         DataField       =   "COM225"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   72
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text66 
         DataField       =   "COM226"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   73
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text67 
         DataField       =   "STA226"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   74
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text68 
         DataField       =   "COM229"
         DataSource      =   "Adodc8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66120
         TabIndex        =   75
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "COM112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   2
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "COM113"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   3
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         DataField       =   "STA111"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   4
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "STA112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   5
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         DataField       =   "MATH111"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   6
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         DataField       =   "MATH112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   7
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         DataField       =   "OTM112"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   8
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         DataField       =   "GNS127"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   9
         Top             =   4200
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   -75000
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         TabIndex        =   189
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
         Left            =   -71760
         TabIndex        =   188
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
         TabIndex        =   187
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
         TabIndex        =   186
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
         TabIndex        =   185
         Top             =   480
         Width           =   5625
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
         Left            =   3240
         TabIndex        =   184
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
         TabIndex        =   183
         Top             =   480
         Width           =   5700
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
         TabIndex        =   182
         Top             =   480
         Width           =   5325
      End
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmCourseNDEdit.frx":24D37
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
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
End
Attribute VB_Name = "frmCourseNDEdit"
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

Private Sub cmdOK_Click()
GetConnect1
Adodc1.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 0
End Sub

Private Sub cmdOK1_Click()
GetConnect2
Adodc2.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 1
End Sub

Private Sub cmdOK2_Click()
GetConnect3
Adodc3.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 2
End Sub

Private Sub cmdOK3_Click()
GetConnect4
Adodc4.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 3
End Sub

Private Sub cmdOK4_Click()
GetConnect5
Adodc5.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 4
End Sub

Private Sub cmdOK5_Click()
GetConnect6
Adodc6.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 5
End Sub

Private Sub cmdOK6_Click()
GetConnect7
Adodc7.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 6
End Sub

Private Sub cmdOK7_Click()
GetConnect8
Adodc8.Recordset.Update
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 7
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
