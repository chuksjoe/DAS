VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNonAcadStaffPI 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Non-Academic Staff Personal Information"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   Icon            =   "frmNonAcadStaffPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNonAcadStaffPI.frx":234CD
   ScaleHeight     =   10170
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNonASPIEdit 
      BackColor       =   &H0080FF80&
      Height          =   10335
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtTrackID 
         DataField       =   "TrackID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   0
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Record"
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   9480
         Width           =   1455
      End
      Begin VB.CommandButton cmdBk2View 
         Caption         =   "&Back"
         Height          =   495
         Left            =   4800
         TabIndex        =   15
         Top             =   9480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddR 
         Caption         =   "&Add New Record"
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   9480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   "Names"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         DataField       =   "Degrees"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         DataField       =   "Position"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         DataField       =   "DateofBirth"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         DataField       =   "Religion"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         DataField       =   "LGA"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   6600
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         DataField       =   "ContactAddress"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   7080
         Width           =   4455
      End
      Begin VB.TextBox Text10 
         DataField       =   "Phone1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   8040
         Width           =   3015
      End
      Begin VB.TextBox Text11 
         DataField       =   "Phone2"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   8520
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Sex"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "frmNonAcadStaffPI.frx":2380F
         Left            =   2400
         List            =   "frmNonAcadStaffPI.frx":23819
         TabIndex        =   2
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cboState 
         DataField       =   "State"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "frmNonAcadStaffPI.frx":2382B
         Left            =   2400
         List            =   "frmNonAcadStaffPI.frx":238A1
         TabIndex        =   8
         Top             =   6120
         Width           =   2415
      End
      Begin VB.ComboBox cboCountry 
         DataField       =   "Country"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "frmNonAcadStaffPI.frx":239C5
         Left            =   2400
         List            =   "frmNonAcadStaffPI.frx":239CF
         TabIndex        =   7
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Track ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   8040
         Width           =   1560
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   4200
         Width           =   1125
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   7080
         Width           =   2235
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.G.A:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   6600
         Width           =   840
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State of Origin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   6120
         Width           =   1950
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   5640
         Width           =   1125
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   5160
         Width           =   1125
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   4680
         Width           =   1680
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NON ACADEMIC STAFF PERSONAL INFO"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Width           =   6645
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   0
         Picture         =   "frmNonAcadStaffPI.frx":239E4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   8520
         Width           =   1560
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.Frame fraNonASPI 
      BackColor       =   &H0000FF00&
      Height          =   10215
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New Record"
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Top             =   9600
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   3840
         TabIndex        =   19
         Top             =   9600
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   495
         Left            =   5160
         TabIndex        =   18
         Top             =   9600
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Top             =   9600
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1080
         Top             =   9120
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
         MaxRecords      =   500
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=StaffSource"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "StaffSource"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblNonAcadstaff"
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
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   0
         Picture         =   "frmNonAcadStaffPI.frx":477B2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NON-ACADEMIC STAFF PERSONAL INFO"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   6675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   4680
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   5160
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   5640
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State of Origin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   6120
         Width           =   1950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.G.A:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   6600
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   7080
         Width           =   2235
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   4200
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   8040
         Width           =   1395
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         DataField       =   "Phone1"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   8040
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         DataField       =   "Names"
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
         Height          =   375
         Left            =   1800
         TabIndex        =   35
         Top             =   2760
         Width           =   5055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         DataField       =   "Degrees"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         DataField       =   "Position"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         DataField       =   "DateofBirth"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         DataField       =   "Religion"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         DataField       =   "Country"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   30
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         DataField       =   "State"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         DataField       =   "LGA"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   6600
         Width           =   2775
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         DataField       =   "ContactAddress"
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
         Height          =   855
         Left            =   2400
         TabIndex        =   27
         Top             =   7080
         Width           =   4455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   8520
         Width           =   1395
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         DataField       =   "Phone2"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   8520
         Width           =   2655
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         DataField       =   "Sex"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Track ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label lblTrackID 
         BackStyle       =   0  'Transparent
         DataField       =   "TrackID"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   2280
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmNonAcadStaffPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect1()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ProjectStaff.mdb;Persist Security Info=False"
End Function

Private Sub cmdAdd_Click()
fraNonASPI.Visible = False
fraNonASPIEdit.Visible = True
End Sub

Private Sub cmdAddR_Click()
On Error GoTo joe
GetConnect1
Adodc1.Recordset.AddNew
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmHome.Show
frmHome.fraHome.Visible = False
frmHome.fraStaff.Visible = True
End Sub

Private Sub cmdBk2View_Click()
fraNonASPI.Visible = True
fraNonASPIEdit.Visible = False
End Sub

Private Sub cmdDelete_Click()
On Error GoTo joe
GetConnect1
don2 = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, "WARNING")
If don2 = vbYes Then
With Adodc1.Recordset
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
con = InputBox("Enter Staff Tracking ID:", "Search by TRACK ID")
Adodc1.Refresh
BookMark1 = Adodc1.Recordset.Bookmark
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "TrackID >= '" & con & "'", 0, adSearchForward
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.Bookmark = BookMark1
MsgBox ("No Record Found")
End If
Exit Sub
joe:
MsgBox Err.Description
End Sub


Private Sub cmdUpdate_Click()
On Error GoTo joe
Adodc1.Recordset.Update
Exit Sub
joe:
MsgBox Err.Description
End Sub
