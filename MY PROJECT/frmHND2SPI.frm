VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHND2SPI 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HND II, Student Personal Information"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930
   Icon            =   "frmHND2SPI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHND2MSEditor 
      BackColor       =   &H0000FF00&
      Height          =   11055
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Frame fraBiaData 
         BackColor       =   &H0080C0FF&
         Caption         =   "Bio Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   98
         Top             =   3360
         Width           =   9735
         Begin VB.Frame Frame1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Names "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   9495
            Begin VB.TextBox TxtFirstname 
               Height          =   285
               Left            =   4680
               TabIndex        =   121
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtSurname 
               Height          =   285
               Left            =   1080
               TabIndex        =   120
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtMiddleI 
               Height          =   285
               Left            =   8400
               MaxLength       =   1
               TabIndex        =   119
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Surname:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   124
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Firsth Name:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3480
               TabIndex        =   123
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Initial:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   7200
               TabIndex        =   122
               Top             =   360
               Width           =   1155
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Date of Birth "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   111
            Top             =   1320
            Width           =   4935
            Begin VB.ComboBox cboDay 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":234CD
               Left            =   600
               List            =   "frmHND2SPI.frx":2352E
               TabIndex        =   114
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox cboMonth 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":235A5
               Left            =   2160
               List            =   "frmHND2SPI.frx":235D0
               TabIndex        =   113
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox cboYear 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":23619
               Left            =   3840
               List            =   "frmHND2SPI.frx":236B6
               TabIndex        =   112
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Day:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   117
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Month:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1440
               TabIndex        =   116
               Top             =   240
               Width           =   585
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Year:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3240
               TabIndex        =   115
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            Caption         =   "Nationality "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   102
            Top             =   2160
            Width           =   9495
            Begin VB.TextBox txtLGA 
               Height          =   285
               Left            =   720
               TabIndex        =   106
               Top             =   720
               Width           =   2415
            End
            Begin VB.ComboBox cboCountry 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":237EC
               Left            =   840
               List            =   "frmHND2SPI.frx":237F6
               TabIndex        =   105
               Top             =   360
               Width           =   2295
            End
            Begin VB.ComboBox cboState 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":2380B
               Left            =   5880
               List            =   "frmHND2SPI.frx":23881
               TabIndex        =   104
               Top             =   360
               Width           =   2055
            End
            Begin VB.ComboBox cboReligion 
               Height          =   315
               ItemData        =   "frmHND2SPI.frx":239A5
               Left            =   5400
               List            =   "frmHND2SPI.frx":239B2
               TabIndex        =   103
               Top             =   720
               Width           =   2535
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Country:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   110
               Top             =   360
               Width           =   720
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "State of Origin:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4560
               TabIndex        =   109
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "L.G.A:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   108
               Top             =   720
               Width           =   525
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Religion:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4560
               TabIndex        =   107
               Top             =   720
               Width           =   795
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0080C0FF&
            Caption         =   "Sex "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5160
            TabIndex        =   99
            Top             =   1320
            Width           =   4455
            Begin VB.OptionButton optMale 
               BackColor       =   &H0080C0FF&
               Caption         =   "Male"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   600
               TabIndex        =   101
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optFemale 
               BackColor       =   &H0080C0FF&
               Caption         =   "Female"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   100
               Top             =   240
               Width           =   1455
            End
         End
      End
      Begin VB.TextBox txtRegNo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   97
         Top             =   2760
         Width           =   615
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Contact details "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   90
         Top             =   6960
         Width           =   4455
         Begin VB.TextBox txtHome 
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox txtContact 
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox txtPhone 
            Height          =   375
            Left            =   1320
            TabIndex        =   91
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Home Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   96
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   95
            Top             =   1320
            Width           =   1530
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   94
            Top             =   2280
            Width           =   945
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Parent/Guardian Details "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   4560
         TabIndex        =   81
         Top             =   6960
         Width           =   5295
         Begin VB.TextBox txtParentN 
            Height          =   315
            Left            =   120
            TabIndex        =   85
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox TxtParentC 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   84
            Top             =   1440
            Width           =   4815
         End
         Begin VB.TextBox txtParentP 
            Height          =   375
            Left            =   1440
            TabIndex        =   83
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtParentOc 
            Height          =   375
            Left            =   1440
            TabIndex        =   82
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parent/Guardian Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   89
            Top             =   480
            Width           =   2115
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   88
            Top             =   1200
            Width           =   1530
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No(s):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   87
            Top             =   2160
            Width           =   1170
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   86
            Top             =   2640
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Height          =   495
         Left            =   5040
         TabIndex        =   80
         Top             =   10440
         Width           =   1575
      End
      Begin VB.CommandButton cmdBk2Pre 
         Caption         =   "&Back"
         Height          =   495
         Left            =   2760
         TabIndex        =   79
         Top             =   10440
         Width           =   1455
      End
      Begin VB.TextBox txtRegNo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   78
         Top             =   2760
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   2055
         Left            =   0
         Picture         =   "frmHND2SPI.frx":239D1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9975
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND II STUDENT PERSONAL INFO EDITOR"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1920
         TabIndex        =   127
         Top             =   2040
         Width           =   6300
      End
      Begin VB.Label Label56 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   126
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbldept 
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   125
         Top             =   2760
         Width           =   615
      End
   End
   Begin VB.Frame fraHND2MSPI 
      BackColor       =   &H0000C000&
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   4200
         TabIndex        =   3
         Top             =   10320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   5520
         TabIndex        =   2
         Top             =   10320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New Record"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   10320
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc adoHND2 
         Height          =   375
         Left            =   2040
         Top             =   9840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
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
         RecordSource    =   "tblHND2StudentDetails"
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
      Begin VB.Line Line2 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   7
         X1              =   120
         X2              =   8760
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line1 
         BorderWidth     =   7
         Index           =   0
         X1              =   120
         X2              =   8760
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HND II STUDENTS PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   37
         Top             =   1800
         Width           =   4950
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Names:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   1950
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   33
         Top             =   4080
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   32
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State of Origin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   31
         Top             =   4800
         Width           =   1845
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.G.A:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   5160
         Width           =   780
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "STUDENT CONTACT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BIO DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1440
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PARENT/GUARDIAN DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   7800
         Width           =   3135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Names:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   8280
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   8640
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   9000
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   9360
         Width           =   1575
      End
      Begin VB.Label lblRegNo 
         BackStyle       =   0  'Transparent
         DataField       =   "RegNo"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   19
         Top             =   2400
         Width           =   1890
      End
      Begin VB.Label lblStuName 
         BackStyle       =   0  'Transparent
         DataField       =   "Student Name"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   18
         Top             =   3360
         Width           =   4770
      End
      Begin VB.Label lbldobirth 
         BackStyle       =   0  'Transparent
         DataField       =   "Date of Birth"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   17
         Top             =   3720
         Width           =   1890
      End
      Begin VB.Label lblSex 
         BackStyle       =   0  'Transparent
         DataField       =   "Sex"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   16
         Top             =   4080
         Width           =   1170
      End
      Begin VB.Label lblCountry 
         BackStyle       =   0  'Transparent
         DataField       =   "Country"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   15
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label lblStaOfOrigin 
         BackStyle       =   0  'Transparent
         DataField       =   "State of Origin"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   14
         Top             =   4800
         Width           =   2370
      End
      Begin VB.Label lblLga 
         BackStyle       =   0  'Transparent
         DataField       =   "L G A"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   13
         Top             =   5160
         Width           =   2130
      End
      Begin VB.Label lblHome 
         BackStyle       =   0  'Transparent
         DataField       =   "Home Address"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   6360
         Width           =   6615
      End
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         DataField       =   "Contact address"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   6720
         Width           =   6615
      End
      Begin VB.Label lblPhone 
         BackStyle       =   0  'Transparent
         DataField       =   "Phone No"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   7080
         Width           =   3015
      End
      Begin VB.Label lblParentN 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Name"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   8280
         Width           =   6015
      End
      Begin VB.Label lblParentC 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Addr"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   8640
         Width           =   6615
      End
      Begin VB.Label lblParentP 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Phone No"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   9000
         Width           =   3135
      End
      Begin VB.Label lblParentOc 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Occupation"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   9360
         Width           =   4335
      End
      Begin VB.Label lblReligion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         DataField       =   "Religion"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         TabIndex        =   5
         Top             =   4440
         Width           =   90
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   4
         Top             =   4440
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   0
         Picture         =   "frmHND2SPI.frx":4779F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8895
      End
   End
   Begin VB.Frame fraHND2MSPreview 
      BackColor       =   &H0080FF80&
      Height          =   10575
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton cmdRedit 
         Caption         =   "&Re-Edit"
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Top             =   9960
         Width           =   1215
      End
      Begin VB.CommandButton cmdBk2SPI 
         Caption         =   "&Back"
         Height          =   375
         Left            =   6480
         TabIndex        =   41
         Top             =   9960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddR 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4920
         TabIndex        =   40
         Top             =   9960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editor"
         Height          =   375
         Left            =   3360
         TabIndex        =   39
         Top             =   9960
         Width           =   1215
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   76
         Top             =   3720
         Width           =   1065
      End
      Begin VB.Label lblParentOcp 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Occupation"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   75
         Top             =   8760
         Width           =   4335
      End
      Begin VB.Label lblParentPp 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Phone No"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   74
         Top             =   8280
         Width           =   3135
      End
      Begin VB.Label lblParentCp 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Addr"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   73
         Top             =   7800
         Width           =   6975
      End
      Begin VB.Label lblParentNp 
         BackStyle       =   0  'Transparent
         DataField       =   "Parent Name"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   72
         Top             =   7320
         Width           =   6015
      End
      Begin VB.Label lblPhonep 
         BackStyle       =   0  'Transparent
         DataField       =   "Phone No"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   71
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Label lblContactp 
         BackStyle       =   0  'Transparent
         DataField       =   "Contact address"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   70
         Top             =   6000
         Width           =   7095
      End
      Begin VB.Label lblHomep 
         BackStyle       =   0  'Transparent
         DataField       =   "Home Address"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   69
         Top             =   5520
         Width           =   7095
      End
      Begin VB.Label lblLgap 
         BackStyle       =   0  'Transparent
         DataField       =   "L G A"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   68
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Label lblStaOfOriginp 
         BackStyle       =   0  'Transparent
         DataField       =   "State of Origin"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   67
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label lblCountryp 
         BackStyle       =   0  'Transparent
         DataField       =   "Country"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   66
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label lblSexp 
         BackStyle       =   0  'Transparent
         DataField       =   "Sex"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   65
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbldobirthp 
         BackStyle       =   0  'Transparent
         DataField       =   "Date of Birth"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   64
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label lblStuNamep 
         BackStyle       =   0  'Transparent
         DataField       =   "Student Name"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   63
         Top             =   2280
         Width           =   5895
      End
      Begin VB.Label lblRegNop 
         BackStyle       =   0  'Transparent
         DataField       =   "RegNo"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   62
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   61
         Top             =   8760
         Width           =   1455
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   60
         Top             =   8280
         Width           =   1260
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   59
         Top             =   7800
         Width           =   2100
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Names:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   58
         Top             =   7320
         Width           =   915
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENT/GUARDIAN DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   57
         Top             =   6840
         Width           =   4275
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BIO DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STUDENT CONTACT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   55
         Top             =   5040
         Width           =   3030
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   6480
         Width           =   1860
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   53
         Top             =   6000
         Width           =   2100
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   52
         Top             =   5520
         Width           =   1860
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.G.A:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   4680
         Width           =   780
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State of Origin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   50
         Top             =   4200
         Width           =   1845
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   49
         Top             =   3720
         Width           =   1020
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   48
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Names:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   46
         Top             =   2280
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STUDENT DETAILS PREVIEW"
         BeginProperty Font 
            Name            =   "Bazooka"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2400
         TabIndex        =   44
         Top             =   720
         Width           =   4560
      End
      Begin VB.Label lblreligionp 
         BackStyle       =   0  'Transparent
         DataField       =   "Religion"
         DataSource      =   "adoHND2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   43
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   0
         Picture         =   "frmHND2SPI.frx":6B56D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmHND2SPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect()
adoHND2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStudents.mdb;Persist Security Info=False"
End Function

Private Sub cmdAdd_Click()
fraHND2MSPI.Visible = False
fraHND2MSPreview.Visible = True
Me.Width = 9570
Me.Height = 10950
End Sub

Private Sub cmdAddR_Click()
On Error GoTo joe
GetConnect
adoHND2.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDM.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab1.Tab = 1
End Sub

Private Sub cmdBk2Pre_Click()
fraHND2MSEditor.Visible = False
fraHND2MSPreview.Visible = True
Me.Width = 9570
Me.Height = 10950
End Sub

Private Sub cmdBk2SPI_Click()
fraHND2MSPI.Visible = True
fraHND2MSPreview.Visible = False
Me.Width = 8955
Me.Height = 11325
End Sub

Private Sub cmdDelete_Click()
On Error GoTo joe
GetConnect
don2 = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, "WARNING")
If don2 = vbYes Then
With adoHND2.Recordset
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

Private Sub cmdEdit_Click()
fraHND2MSEditor.Visible = True
fraHND2MSPreview.Visible = False
Me.Width = 10080
Me.Height = 11520
txtSurname = ""
TxtFirstname = ""
txtMiddleI = ""
cboDay = ""
cboMonth = ""
cboYear = ""
cboCountry = ""
cboState = ""
cboReligion = ""
txtLGA = ""
txtContact = ""
txtHome = ""
TxtParentC = ""
txtParentN = ""
txtParentP = ""
txtParentOc = ""
txtPhone = ""
txtRegNo1 = ""
txtRegNo2 = ""
txtRegNo1.SetFocus
End Sub

Private Sub cmdPreview_Click()
fraHND2MSEditor.Visible = False
fraHND2MSPreview.Visible = True
Me.Width = 9570
Me.Height = 10950
lblRegNop.Caption = txtRegNo1.Text + "/" + txtRegNo2.Text + "/" + lbldept.Caption
lblStuNamep.Caption = txtSurname.Text + " " + TxtFirstname.Text + " " + txtMiddleI.Text
lbldobirthp.Caption = cboDay.Text + "/" + cboMonth.Text + "/" + cboYear.Text
If optMale.Value Then
lblSexp = "Male"
ElseIf optFemale.Value Then
lblSexp = "Female"
End If
lblCountryp.Caption = cboCountry.Text
lblStaOfOriginp.Caption = cboState.Text
lblreligionp.Caption = cboReligion.Text
lblLgap.Caption = txtLGA.Text
lblContactp.Caption = txtContact.Text
lblHomep.Caption = txtHome.Text
lblPhonep.Caption = txtPhone.Text
lblParentNp.Caption = txtParentN.Text
lblParentCp.Caption = TxtParentC.Text
lblParentPp.Caption = txtParentP.Text
lblParentOcp.Caption = txtParentOc.Text
End Sub

Private Sub cmdRedit_Click()
fraHND2MSEditor.Visible = True
fraHND2MSPreview.Visible = False
Me.Width = 10080
Me.Height = 11520
End Sub


