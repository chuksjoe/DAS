VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmND1EExam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ND I(Evening) Examination Details"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   Icon            =   "frmND1EExam.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraND1EEEdit 
      BackColor       =   &H0080FFFF&
      Height          =   10215
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      Begin TabDlg.SSTab SSTab2 
         Height          =   8415
         Left            =   0
         TabIndex        =   98
         Top             =   1800
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   14843
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "First Semester"
         TabPicture(0)   =   "frmND1EExam.frx":234CD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdCompute"
         Tab(0).Control(1)=   "cmdAddR"
         Tab(0).Control(2)=   "cmdBk2Exam1"
         Tab(0).Control(3)=   "Combo8"
         Tab(0).Control(4)=   "Combo7"
         Tab(0).Control(5)=   "txtName"
         Tab(0).Control(6)=   "Combo1"
         Tab(0).Control(7)=   "Combo6"
         Tab(0).Control(8)=   "Combo5"
         Tab(0).Control(9)=   "Combo4"
         Tab(0).Control(10)=   "Combo3"
         Tab(0).Control(11)=   "Combo2"
         Tab(0).Control(12)=   "txtRegNo"
         Tab(0).Control(13)=   "Combo9"
         Tab(0).Control(14)=   "lblTotal"
         Tab(0).Control(15)=   "lblGpa"
         Tab(0).Control(16)=   "Label50"
         Tab(0).Control(17)=   "Label49"
         Tab(0).Control(18)=   "Label48"
         Tab(0).Control(19)=   "Label47"
         Tab(0).Control(20)=   "Label45"
         Tab(0).Control(21)=   "Label44"
         Tab(0).Control(22)=   "Label34"
         Tab(0).Control(23)=   "Label28"
         Tab(0).Control(24)=   "Label27"
         Tab(0).Control(25)=   "Label24"
         Tab(0).Control(26)=   "Label19"
         Tab(0).Control(27)=   "Label18"
         Tab(0).Control(28)=   "Label17"
         Tab(0).Control(29)=   "Label15"
         Tab(0).ControlCount=   30
         TabCaption(1)   =   "Second Semester"
         TabPicture(1)   =   "frmND1EExam.frx":234E9
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label53"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label54"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label55"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label56"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label57"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label58"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label59"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label60"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label61"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label62"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label63"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label64"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label65"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label66"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label67"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "lblGpa2"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "lblTotal2"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Combo10"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Text1"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Combo11"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Combo12"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "Combo13"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Combo14"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "Combo15"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "Combo16"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "Text2"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "Combo17"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "Combo18"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "Combo19"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "cmdCompute2"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "cmdAddR2"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "cmdBK2Exam2"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).ControlCount=   32
         Begin VB.CommandButton cmdBK2Exam2 
            Caption         =   "&Back"
            Height          =   495
            Left            =   4680
            TabIndex        =   28
            Top             =   7200
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddR2 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   3360
            TabIndex        =   27
            Top             =   7200
            Width           =   1215
         End
         Begin VB.CommandButton cmdCompute2 
            Caption         =   "&Commpute GPA"
            Height          =   495
            Left            =   1680
            TabIndex        =   26
            Top             =   7200
            Width           =   1575
         End
         Begin VB.ComboBox Combo19 
            DataField       =   "URP120"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23505
            Left            =   2520
            List            =   "frmND1EExam.frx":23524
            TabIndex        =   25
            Text            =   "..."
            Top             =   6480
            Width           =   1215
         End
         Begin VB.ComboBox Combo18 
            DataField       =   "EED126"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23546
            Left            =   2520
            List            =   "frmND1EExam.frx":23565
            TabIndex        =   24
            Text            =   "..."
            Top             =   6120
            Width           =   1215
         End
         Begin VB.ComboBox Combo17 
            DataField       =   "GNS128"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23587
            Left            =   2520
            List            =   "frmND1EExam.frx":235A6
            TabIndex        =   23
            Text            =   "..."
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            DataField       =   "Names"
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   1935
            TabIndex        =   15
            Top             =   2160
            Width           =   5175
         End
         Begin VB.ComboBox Combo16 
            DataField       =   "GNS102"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":235C8
            Left            =   2520
            List            =   "frmND1EExam.frx":235E7
            TabIndex        =   22
            Text            =   "..."
            Top             =   5400
            Width           =   1215
         End
         Begin VB.ComboBox Combo15 
            DataField       =   "COM126"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23609
            Left            =   2520
            List            =   "frmND1EExam.frx":23628
            TabIndex        =   21
            Text            =   "..."
            Top             =   5040
            Width           =   1215
         End
         Begin VB.ComboBox Combo14 
            DataField       =   "COM125"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":2364A
            Left            =   2520
            List            =   "frmND1EExam.frx":23669
            TabIndex        =   20
            Text            =   "..."
            Top             =   4680
            Width           =   1215
         End
         Begin VB.ComboBox Combo13 
            DataField       =   "COM124"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":2368B
            Left            =   2520
            List            =   "frmND1EExam.frx":236AA
            TabIndex        =   19
            Text            =   "..."
            Top             =   4320
            Width           =   1215
         End
         Begin VB.ComboBox Combo12 
            DataField       =   "COM123"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":236CC
            Left            =   2520
            List            =   "frmND1EExam.frx":236EB
            TabIndex        =   18
            Text            =   "..."
            Top             =   3960
            Width           =   1215
         End
         Begin VB.ComboBox Combo11 
            DataField       =   "COM122"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":2370D
            Left            =   2520
            List            =   "frmND1EExam.frx":2372C
            TabIndex        =   17
            Text            =   "..."
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox Text1 
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
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   1920
            TabIndex        =   14
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox Combo10 
            DataField       =   "COM121"
            DataSource      =   "adoND1Eexam2s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":2374E
            Left            =   2520
            List            =   "frmND1EExam.frx":2376D
            TabIndex        =   16
            Text            =   "..."
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCompute 
            Caption         =   "&Commpute GPA"
            Height          =   495
            Left            =   -73320
            TabIndex        =   11
            Top             =   6960
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddR 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   -71640
            TabIndex        =   12
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdBk2Exam1 
            Caption         =   "&Back"
            Height          =   495
            Left            =   -70320
            TabIndex        =   13
            Top             =   6960
            Width           =   1215
         End
         Begin VB.ComboBox Combo8 
            DataField       =   "OTM112"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":2378F
            Left            =   -72480
            List            =   "frmND1EExam.frx":237AE
            TabIndex        =   9
            Text            =   "..."
            Top             =   5760
            Width           =   1215
         End
         Begin VB.ComboBox Combo7 
            DataField       =   "MATH112"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":237D0
            Left            =   -72480
            List            =   "frmND1EExam.frx":237EF
            TabIndex        =   8
            Text            =   "..."
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            DataField       =   "Names"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -73065
            TabIndex        =   1
            Top             =   2160
            Width           =   5175
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "COM101"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23811
            Left            =   -72480
            List            =   "frmND1EExam.frx":23830
            TabIndex        =   2
            Text            =   "..."
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ComboBox Combo6 
            DataField       =   "MATH111"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23852
            Left            =   -72480
            List            =   "frmND1EExam.frx":23871
            TabIndex        =   7
            Text            =   "..."
            Top             =   5040
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            DataField       =   "STA112"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23893
            Left            =   -72480
            List            =   "frmND1EExam.frx":238B2
            TabIndex        =   6
            Text            =   "..."
            Top             =   4680
            Width           =   1215
         End
         Begin VB.ComboBox Combo4 
            DataField       =   "STA111"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":238D4
            Left            =   -72480
            List            =   "frmND1EExam.frx":238F3
            TabIndex        =   5
            Text            =   "..."
            Top             =   4320
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "COM113"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23915
            Left            =   -72480
            List            =   "frmND1EExam.frx":23934
            TabIndex        =   4
            Text            =   "..."
            Top             =   3960
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "COM112"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23956
            Left            =   -72480
            List            =   "frmND1EExam.frx":23975
            TabIndex        =   3
            Text            =   "..."
            Top             =   3600
            Width           =   1215
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
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -73080
            TabIndex        =   0
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox Combo9 
            DataField       =   "GNS127"
            DataSource      =   "adoND1Eexam1s"
            Height          =   315
            ItemData        =   "frmND1EExam.frx":23997
            Left            =   -72480
            List            =   "frmND1EExam.frx":239B6
            TabIndex        =   10
            Text            =   "..."
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label lblTotal2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   5160
            TabIndex        =   131
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label lblGpa2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   5160
            TabIndex        =   130
            Top             =   5040
            Width           =   1935
         End
         Begin VB.Label Label67 
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
            Left            =   4320
            TabIndex        =   129
            Top             =   5040
            Width           =   720
         End
         Begin VB.Label Label66 
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
            Left            =   4320
            TabIndex        =   128
            Top             =   4440
            Width           =   765
         End
         Begin VB.Label Label65 
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
            Left            =   720
            TabIndex        =   127
            Top             =   1560
            Width           =   1080
         End
         Begin VB.Label Label64 
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
            Left            =   945
            TabIndex        =   126
            Top             =   2160
            Width           =   870
         End
         Begin VB.Label Label63 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ND I(EVENING) EXAMINATION DETAILS (SECOND SEMESTER)"
            BeginProperty Font 
               Name            =   "Amelia"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   1320
            TabIndex        =   125
            Top             =   720
            Width           =   5205
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS102 -"
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
            Left            =   960
            TabIndex        =   124
            Top             =   5400
            Width           =   1275
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM121 -"
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
            Left            =   960
            TabIndex        =   123
            Top             =   3240
            Width           =   1350
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM122 -"
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
            Left            =   960
            TabIndex        =   122
            Top             =   3600
            Width           =   1350
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM123 -"
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
            Left            =   960
            TabIndex        =   121
            Top             =   3960
            Width           =   1350
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM124 -"
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
            Left            =   960
            TabIndex        =   120
            Top             =   4320
            Width           =   1350
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM125 -"
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
            Left            =   960
            TabIndex        =   119
            Top             =   4680
            Width           =   1350
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM126 -"
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
            Left            =   960
            TabIndex        =   118
            Top             =   5040
            Width           =   1350
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EED126 -"
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
            Left            =   960
            TabIndex        =   117
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS128 -"
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
            Left            =   960
            TabIndex        =   116
            Top             =   5760
            Width           =   1275
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URP120 -"
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
            Left            =   960
            TabIndex        =   115
            Top             =   6480
            Width           =   1230
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -69840
            TabIndex        =   114
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label lblGpa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -69840
            TabIndex        =   113
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MATH112-"
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
            Left            =   -74040
            TabIndex        =   112
            Top             =   5400
            Width           =   1485
         End
         Begin VB.Label Label49 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ND I(EVENING) EXAMINATION DETAILS (FIRST SEMESTER)"
            BeginProperty Font 
               Name            =   "Amelia"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   -73680
            TabIndex        =   111
            Top             =   720
            Width           =   5205
         End
         Begin VB.Label Label48 
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
            Left            =   -74055
            TabIndex        =   110
            Top             =   2160
            Width           =   870
         End
         Begin VB.Label Label47 
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
            Left            =   -70680
            TabIndex        =   109
            Top             =   4800
            Width           =   720
         End
         Begin VB.Label Label45 
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
            Left            =   -70680
            TabIndex        =   108
            Top             =   4200
            Width           =   765
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM101 -"
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
            Left            =   -74040
            TabIndex        =   107
            Top             =   3240
            Width           =   1350
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM112 -"
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
            Left            =   -74040
            TabIndex        =   106
            Top             =   3600
            Width           =   1350
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM113 -"
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
            Left            =   -74040
            TabIndex        =   105
            Top             =   3960
            Width           =   1350
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA111 -"
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
            Left            =   -74040
            TabIndex        =   104
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA112 -"
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
            Left            =   -74040
            TabIndex        =   103
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MATH111-"
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
            Left            =   -74040
            TabIndex        =   102
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label18 
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
            Left            =   -74280
            TabIndex        =   101
            Top             =   1560
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS127 -"
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
            Left            =   -74040
            TabIndex        =   100
            Top             =   6120
            Width           =   1275
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS101 -"
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
            Left            =   -74040
            TabIndex        =   99
            Top             =   5760
            Width           =   1275
         End
      End
      Begin VB.Image Image2 
         Height          =   1815
         Left            =   0
         Picture         =   "frmND1EExam.frx":239D8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.Frame fraND1EExam 
      BackColor       =   &H00C0E0FF&
      Height          =   10215
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8175
      Begin TabDlg.SSTab SSTab1 
         Height          =   8535
         Left            =   0
         TabIndex        =   30
         Top             =   1800
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   15055
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12648447
         TabCaption(0)   =   "First Semester"
         TabPicture(0)   =   "frmND1EExam.frx":477A6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdAdd"
         Tab(0).Control(1)=   "cmdDelete"
         Tab(0).Control(2)=   "cmdSearch"
         Tab(0).Control(3)=   "cmdBack"
         Tab(0).Control(4)=   "adoND1Eexam1s"
         Tab(0).Control(5)=   "Label41"
         Tab(0).Control(6)=   "Label40"
         Tab(0).Control(7)=   "Label39"
         Tab(0).Control(8)=   "Label38"
         Tab(0).Control(9)=   "Label14"
         Tab(0).Control(10)=   "Label13"
         Tab(0).Control(11)=   "Label3"
         Tab(0).Control(12)=   "Label4"
         Tab(0).Control(13)=   "Label5"
         Tab(0).Control(14)=   "Label6"
         Tab(0).Control(15)=   "Label7"
         Tab(0).Control(16)=   "Label8"
         Tab(0).Control(17)=   "Label9"
         Tab(0).Control(18)=   "Label10"
         Tab(0).Control(19)=   "Gpa"
         Tab(0).Control(20)=   "Total"
         Tab(0).Control(21)=   "Label11"
         Tab(0).Control(22)=   "Label1"
         Tab(0).Control(23)=   "lblCom412"
         Tab(0).Control(24)=   "lblSta411"
         Tab(0).Control(25)=   "lblCom416"
         Tab(0).Control(26)=   "lblCom415"
         Tab(0).Control(27)=   "lblCom414"
         Tab(0).Control(28)=   "lblCom413"
         Tab(0).Control(29)=   "lblName"
         Tab(0).Control(30)=   "lblRegNo"
         Tab(0).Control(31)=   "Line1"
         Tab(0).Control(32)=   "Label33"
         Tab(0).Control(33)=   "Label35"
         Tab(0).ControlCount=   34
         TabCaption(1)   =   "Second Semester"
         TabPicture(1)   =   "frmND1EExam.frx":477C2
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Total2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Gpa2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label21"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label20"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label2"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label25"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label26"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lblName2"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lblRegNo2"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Line2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lblgns102b"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label16"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "lblcom122b"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "lblcom123b"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "lblcom124b"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "lblcom125b"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "lblcom126b"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "lblcom121b"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Label29"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Label30"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "Label31"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Label32"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "Label36"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "Label37"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "Label42"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "Label43"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "lbleed126b"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "lblgns128b"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "Label46"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "lblurp120b"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "adoND1Eexam2s"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "cmdAdd2"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "cmdDelete2"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "cmdSearch2"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "cmdBack2"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).ControlCount=   36
         Begin VB.CommandButton cmdBack2 
            Caption         =   "&Back"
            Height          =   495
            Left            =   5160
            TabIndex        =   38
            Top             =   7800
            Width           =   1215
         End
         Begin VB.CommandButton cmdSearch2 
            Caption         =   "&Search Rec."
            Height          =   495
            Left            =   3840
            TabIndex        =   37
            Top             =   7800
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete2 
            Caption         =   "&Delete Rec."
            Height          =   495
            Left            =   2520
            TabIndex        =   36
            Top             =   7800
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd2 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   1200
            TabIndex        =   35
            Top             =   7800
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add Record"
            Height          =   495
            Left            =   -73680
            TabIndex        =   34
            Top             =   7440
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete Rec."
            Height          =   495
            Left            =   -72360
            TabIndex        =   33
            Top             =   7440
            Width           =   1215
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search Rec."
            Height          =   495
            Left            =   -71040
            TabIndex        =   32
            Top             =   7440
            Width           =   1215
         End
         Begin VB.CommandButton cmdBack 
            Caption         =   "&Back"
            Height          =   495
            Left            =   -69720
            TabIndex        =   31
            Top             =   7440
            Width           =   1215
         End
         Begin MSAdodcLib.Adodc adoND1Eexam1s 
            Height          =   615
            Left            =   -74040
            Top             =   6600
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
            Connect         =   "DSN=HND2firstSConc"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "HND2firstSConc"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "tblND1EFirst"
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
         Begin MSAdodcLib.Adodc adoND1Eexam2s 
            Height          =   615
            Left            =   840
            Top             =   6960
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
            Connect         =   "DSN=HND2firstSConc"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "HND2firstSConc"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "tblND1ESecond"
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
         Begin VB.Label lblurp120b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "URP120"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   96
            Top             =   6480
            Width           =   615
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URP120 -"
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
            Left            =   1320
            TabIndex        =   95
            Top             =   6480
            Width           =   1230
         End
         Begin VB.Label lblgns128b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "GNS128"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   94
            Top             =   5760
            Width           =   615
         End
         Begin VB.Label lbleed126b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "EED126"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   93
            Top             =   6120
            Width           =   615
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS128 -"
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
            Left            =   1320
            TabIndex        =   92
            Top             =   5760
            Width           =   1275
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EED126 -"
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
            Left            =   1320
            TabIndex        =   91
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM126 -"
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
            Left            =   1320
            TabIndex        =   90
            Top             =   5040
            Width           =   1350
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM125 -"
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
            Left            =   1320
            TabIndex        =   89
            Top             =   4680
            Width           =   1350
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM124 -"
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
            Left            =   1320
            TabIndex        =   88
            Top             =   4320
            Width           =   1350
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM123 -"
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
            Left            =   1320
            TabIndex        =   87
            Top             =   3960
            Width           =   1350
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM122 -"
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
            Left            =   1320
            TabIndex        =   86
            Top             =   3600
            Width           =   1350
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM121 -"
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
            Left            =   1320
            TabIndex        =   85
            Top             =   3240
            Width           =   1350
         End
         Begin VB.Label lblcom121b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM121"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   84
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label lblcom126b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM126"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   83
            Top             =   5040
            Width           =   615
         End
         Begin VB.Label lblcom125b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM125"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   82
            Top             =   4680
            Width           =   615
         End
         Begin VB.Label lblcom124b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM124"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   81
            Top             =   4320
            Width           =   615
         End
         Begin VB.Label lblcom123b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM123"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   80
            Top             =   3960
            Width           =   615
         End
         Begin VB.Label lblcom122b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "COM122"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   79
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS102 -"
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
            Left            =   1320
            TabIndex        =   78
            Top             =   5400
            Width           =   1275
         End
         Begin VB.Label lblgns102b 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "GNS102"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   77
            Top             =   5400
            Width           =   615
         End
         Begin VB.Label Label41 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "OTM112"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   76
            Top             =   5760
            Width           =   615
         End
         Begin VB.Label Label40 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "GNS127"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   75
            Top             =   6120
            Width           =   615
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS101 -"
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
            Left            =   -73680
            TabIndex        =   74
            Top             =   5760
            Width           =   1275
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GNS127 -"
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
            Left            =   -73680
            TabIndex        =   73
            Top             =   6120
            Width           =   1275
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   120
            X2              =   7560
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label lblRegNo2 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "RegNo"
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   2040
            TabIndex        =   72
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lblName2 
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            DataField       =   "Names"
            DataSource      =   "adoND1Eexam2s"
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
            Left            =   2040
            TabIndex        =   71
            Top             =   2160
            Width           =   5055
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ND I(EVENING) EXAMINATION DETAILS (SECOND SEMESTER)"
            BeginProperty Font 
               Name            =   "Amelia"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   1320
            TabIndex        =   70
            Top             =   720
            Width           =   5205
         End
         Begin VB.Label Label25 
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
            Left            =   945
            TabIndex        =   69
            Top             =   2160
            Width           =   870
         End
         Begin VB.Label Label12 
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
            Left            =   2880
            TabIndex        =   68
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label2 
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
            Left            =   720
            TabIndex        =   67
            Top             =   1560
            Width           =   1080
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
            Left            =   -74280
            TabIndex        =   66
            Top             =   1560
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
            Left            =   -72120
            TabIndex        =   65
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MATH111-"
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
            Left            =   -73680
            TabIndex        =   64
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA112 -"
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
            Left            =   -73680
            TabIndex        =   63
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STA111 -"
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
            Left            =   -73680
            TabIndex        =   62
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM113 -"
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
            Left            =   -73680
            TabIndex        =   61
            Top             =   3960
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM112 -"
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
            Left            =   -73680
            TabIndex        =   60
            Top             =   3600
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COM101 -"
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
            Left            =   -73680
            TabIndex        =   59
            Top             =   3240
            Width           =   1350
         End
         Begin VB.Label Label9 
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
            Left            =   -71040
            TabIndex        =   58
            Top             =   3960
            Width           =   765
         End
         Begin VB.Label Label10 
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
            Left            =   -71040
            TabIndex        =   57
            Top             =   4560
            Width           =   720
         End
         Begin VB.Label Gpa 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -70200
            TabIndex        =   56
            Top             =   4560
            Width           =   225
         End
         Begin VB.Label Total 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -70200
            TabIndex        =   55
            Top             =   3960
            Width           =   225
         End
         Begin VB.Label Label11 
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
            Left            =   -74055
            TabIndex        =   54
            Top             =   2160
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ND I(EVENING) EXAMINATION DETAILS (FIRST SEMESTER)"
            BeginProperty Font 
               Name            =   "Amelia"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   -73680
            TabIndex        =   53
            Top             =   720
            Width           =   5205
         End
         Begin VB.Label lblCom412 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM101"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   52
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label lblSta411 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "MATH111"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   51
            Top             =   5040
            Width           =   615
         End
         Begin VB.Label lblCom416 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "STA112"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   50
            Top             =   4680
            Width           =   615
         End
         Begin VB.Label lblCom415 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "STA111"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   49
            Top             =   4320
            Width           =   615
         End
         Begin VB.Label lblCom414 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM113"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   48
            Top             =   3960
            Width           =   615
         End
         Begin VB.Label lblCom413 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "COM112"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   47
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lblName 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "Names"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72960
            TabIndex        =   46
            Top             =   2160
            Width           =   5055
         End
         Begin VB.Label lblRegNo 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "RegNo"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72960
            TabIndex        =   45
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   -74880
            X2              =   -67440
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MATH112-"
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
            Left            =   -73680
            TabIndex        =   44
            Top             =   5400
            Width           =   1485
         End
         Begin VB.Label Label35 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            DataField       =   "MATH112"
            DataSource      =   "adoND1Eexam1s"
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
            Left            =   -72120
            TabIndex        =   43
            Top             =   5400
            Width           =   615
         End
         Begin VB.Label Label20 
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
            TabIndex        =   42
            Top             =   3960
            Width           =   765
         End
         Begin VB.Label Label21 
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
            TabIndex        =   41
            Top             =   4560
            Width           =   720
         End
         Begin VB.Label Gpa2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "GPA"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   40
            Top             =   4560
            Width           =   225
         End
         Begin VB.Label Total2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            DataField       =   "Total"
            DataSource      =   "adoND1Eexam2s"
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
            TabIndex        =   39
            Top             =   3960
            Width           =   225
         End
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   0
         Picture         =   "frmND1EExam.frx":477DE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmND1EExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const totalchr As Single = 29
Const totalchr2 As Single = 38
Private Function GetConnect1()
adoND1Eexam1s.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function
Private Function GetConnect2()
adoND1Eexam2s.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function

Private Sub cmdAdd_Click()
fraND1EExam.Visible = False
fraND1EEEdit.Visible = True
SSTab2.Tab = 0
End Sub

Private Sub cmdAdd2_Click()
fraND1EExam.Visible = False
fraND1EEEdit.Visible = True
SSTab2.Tab = 1
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmNd.Show
frmNd.fraNDE.Visible = True
frmNd.fraND.Visible = False
frmNd.SSTab1.Tab = 0
End Sub

Private Sub cmdBack2_Click()
Me.Hide
frmNd.Show
frmNd.fraNDE.Visible = True
frmNd.fraND.Visible = False
frmNd.SSTab1.Tab = 0
End Sub

Private Sub cmdBk2Exam1_Click()
fraND1EEEdit.Visible = False
fraND1EExam.Visible = True
SSTab1.Tab = 0
End Sub

Private Sub cmdBk2Exam2_Click()
fraND1EEEdit.Visible = False
fraND1EExam.Visible = True
SSTab1.Tab = 1
End Sub

Private Sub cmdDelete_Click()
On Error GoTo joe
GetConnect1
don = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, "WARNING")
If don = vbYes Then
With adoND1Eexam1s.Recordset
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
With adoND1Eexam2s.Recordset
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
BookMark1 = adoND1Eexam1s.Recordset.Bookmark
adoND1Eexam1s.Recordset.MoveFirst
adoND1Eexam1s.Recordset.Find "regno = '" & con & "'", 0, adSearchForward
If adoND1Eexam1s.Recordset.EOF = True Then
adoND1Eexam1s.Recordset.Bookmark = BookMark1
MsgBox ("No Record Found")
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdAddR_Click()
On Error GoTo joe
GetConnect1
Adodc1.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdAddR2_Click()
On Error GoTo joe
GetConnect2
Adodc2.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdCompute_Click()
Dim val1 As Single, val2 As Single, val3 As Single, val4 As Single, val5 As Single, val6 As Single, val7 As Single
Dim val8 As Single, val9 As Single
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

val1 = grade1 * 4
val2 = grade2 * 4
val3 = grade3 * 4
val4 = grade4 * 3
val5 = grade5 * 3
val6 = grade6 * 2
val7 = grade7 * 3
val8 = grade8 * 4
val9 = grade9 * 2

Total = val1 + val2 + val3 + val4 + val5 + val6 + val7 + val8 + val9
lblTotal.Caption = Total

Gpa = Total / totalchr
lblGpa.Caption = Gpa
End Sub


Private Sub cmdCompute2_Click()
Dim val10 As Single, val11 As Single, val12 As Single, val13 As Single, val14 As Single, val15 As Single, val16 As Single
Dim val17 As Single, val18 As Single, val19 As Single
Dim Total2 As Single
Dim Gpa2 As Single


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
If Combo13.Text = "A" Then
grade13 = 4
ElseIf Combo13.Text = "AB" Then
grade13 = 3.5
ElseIf Combo13.Text = "B" Then
grade13 = 3
ElseIf Combo13.Text = "BC" Then
grade13 = 2.5
ElseIf Combo13.Text = "C" Then
grade13 = 2
ElseIf Combo13.Text = "CD" Then
grade13 = 1.5
ElseIf Combo13.Text = "D" Then
grade13 = 1
ElseIf Combo13.Text = "E" Then
grade13 = 0.5
ElseIf Combo13.Text = "F" Then
grade13 = 0
End If
If Combo14.Text = "A" Then
grade14 = 4
ElseIf Combo14.Text = "AB" Then
grade14 = 3.5
ElseIf Combo14.Text = "B" Then
grade14 = 3
ElseIf Combo14.Text = "BC" Then
grade14 = 2.5
ElseIf Combo14.Text = "C" Then
grade14 = 2
ElseIf Combo14.Text = "CD" Then
grade14 = 1.5
ElseIf Combo14.Text = "D" Then
grade14 = 1
ElseIf Combo14.Text = "E" Then
grade14 = 0.5
ElseIf Combo14.Text = "F" Then
grade14 = 0
End If
If Combo15.Text = "A" Then
grade15 = 4
ElseIf Combo15.Text = "AB" Then
grade15 = 3.5
ElseIf Combo15.Text = "B" Then
grade15 = 3
ElseIf Combo15.Text = "BC" Then
grade15 = 2.5
ElseIf Combo15.Text = "C" Then
grade15 = 2
ElseIf Combo15.Text = "CD" Then
grade15 = 1.5
ElseIf Combo15.Text = "D" Then
grade15 = 1
ElseIf Combo15.Text = "E" Then
grade15 = 0.5
ElseIf Combo15.Text = "F" Then
grade15 = 0
End If
If Combo16.Text = "A" Then
grade16 = 4
ElseIf Combo16.Text = "AB" Then
grade16 = 3.5
ElseIf Combo16.Text = "B" Then
grade16 = 3
ElseIf Combo16.Text = "BC" Then
grade16 = 2.5
ElseIf Combo16.Text = "C" Then
grade16 = 2
ElseIf Combo16.Text = "CD" Then
grade16 = 1.5
ElseIf Combo16.Text = "D" Then
grade16 = 1
ElseIf Combo16.Text = "E" Then
grade16 = 0.5
ElseIf Combo16.Text = "F" Then
grade16 = 0
End If
If Combo17.Text = "A" Then
grade17 = 4
ElseIf Combo17.Text = "AB" Then
grade17 = 3.5
ElseIf Combo17.Text = "B" Then
grade17 = 3
ElseIf Combo17.Text = "BC" Then
grade17 = 2.5
ElseIf Combo17.Text = "C" Then
grade17 = 2
ElseIf Combo17.Text = "CD" Then
grade17 = 1.5
ElseIf Combo17.Text = "D" Then
grade17 = 1
ElseIf Combo17.Text = "E" Then
grade17 = 0.5
ElseIf Combo17.Text = "F" Then
grade17 = 0
End If
If Combo18.Text = "A" Then
grade18 = 4
ElseIf Combo18.Text = "AB" Then
grade18 = 3.5
ElseIf Combo18.Text = "B" Then
grade18 = 3
ElseIf Combo18.Text = "BC" Then
grade18 = 2.5
ElseIf Combo18.Text = "C" Then
grade18 = 2
ElseIf Combo18.Text = "CD" Then
grade18 = 1.5
ElseIf Combo18.Text = "D" Then
grade18 = 1
ElseIf Combo18.Text = "E" Then
grade18 = 0.5
ElseIf Combo18.Text = "F" Then
grade18 = 0
End If
If Combo19.Text = "A" Then
grade19 = 4
ElseIf Combo19.Text = "AB" Then
grade19 = 3.5
ElseIf Combo19.Text = "B" Then
grade19 = 3
ElseIf Combo19.Text = "BC" Then
grade19 = 2.5
ElseIf Combo19.Text = "C" Then
grade19 = 2
ElseIf Combo19.Text = "CD" Then
grade19 = 1.5
ElseIf Combo19.Text = "D" Then
grade19 = 1
ElseIf Combo19.Text = "E" Then
grade19 = 0.5
ElseIf Combo19.Text = "F" Then
grade19 = 0
End If


val10 = grade10 * 6
val11 = grade11 * 4
val12 = grade12 * 6
val13 = grade13 * 4
val14 = grade14 * 3
val15 = grade15 * 6
val16 = grade16 * 2
val17 = grade17 * 3
val18 = grade18 * 3
val19 = grade19 * 1

Total2 = val10 + val11 + val12 + val13 + val14 + val15 + val16 + val17 + val18 + val19
lblTotal2.Caption = Total2

Gpa2 = Total2 / totalchr2
lblGpa2.Caption = Gpa2
End Sub

Private Sub cmdSearch2_Click()
On Error GoTo joe
GetConnect2
Dim con As String
con = InputBox("Enter Student Reg. Number", "Search By Reg. No.")
BookMark1 = adoND1Eexam2s.Recordset.Bookmark
adoND1Eexam2s.Recordset.MoveFirst
adoND1Eexam2s.Recordset.Find "regno = '" & con & "'", 0, adSearchForward
If adoND1Eexam2s.Recordset.EOF = True Then
adoND1Eexam2s.Recordset.Bookmark = BookMark1
MsgBox ("No Record Found")
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub
