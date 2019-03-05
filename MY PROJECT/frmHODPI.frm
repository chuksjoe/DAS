VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHODPI 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HOD Personal Infomation"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930
   Icon            =   "frmHODPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPPHolder 
      Height          =   2655
      Left            =   7200
      TabIndex        =   80
      Top             =   2640
      Width           =   2175
      Begin VB.Image ImaPP 
         Height          =   2715
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2160
      End
   End
   Begin MSComDlg.CommonDialog cd1TesT 
      Left            =   7560
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   7440
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   "DSN=StaffSource"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "StaffSource"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblHOD"
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
   Begin VB.Frame fraView 
      BackColor       =   &H0000FF00&
      Height          =   8655
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   3600
         TabIndex        =   58
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   5160
         TabIndex        =   21
         Top             =   8040
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "BIO-DATA"
         TabPicture(0)   =   "frmHODPI.frx":234CD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lbltag(8)"
         Tab(0).Control(1)=   "lblTitle"
         Tab(0).Control(2)=   "lblName"
         Tab(0).Control(3)=   "lblSex"
         Tab(0).Control(4)=   "lblHonours"
         Tab(0).Control(5)=   "lblDob"
         Tab(0).Control(6)=   "lblReligion"
         Tab(0).Control(7)=   "lblCountry"
         Tab(0).Control(8)=   "lblState"
         Tab(0).Control(9)=   "lblLGA"
         Tab(0).Control(10)=   "lbltag(11)"
         Tab(0).Control(11)=   "lbltag(17)"
         Tab(0).Control(12)=   "lbltag(16)"
         Tab(0).Control(13)=   "lbltag(15)"
         Tab(0).Control(14)=   "lbltag(14)"
         Tab(0).Control(15)=   "lbltag(13)"
         Tab(0).Control(16)=   "lbltag(12)"
         Tab(0).Control(17)=   "lbltag(10)"
         Tab(0).Control(18)=   "lbltag(9)"
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "CONTACT DETAILS"
         TabPicture(1)   =   "frmHODPI.frx":234E9
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lbltag(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbltag(6)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbltag(4)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lbltag(3)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbltag(2)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lbltag(7)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lbltag(5)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lblNofKinPhone"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lblRelation"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lblNofKin"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "lblPhone2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lblPhone1"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "lblAddress"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "lblNofKinAd"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "lbltag(0)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "BIO-DATA."
            BeginProperty Font 
               Name            =   "Angie Groovin"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   8
            Left            =   -72360
            TabIndex        =   56
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Title"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   55
            Top             =   1200
            Width           =   1365
         End
         Begin VB.Label lblName 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Names"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   54
            Top             =   1680
            Width           =   4725
         End
         Begin VB.Label lblSex 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Sex"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   53
            Top             =   2160
            Width           =   1365
         End
         Begin VB.Label lblHonours 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Honours"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   52
            Top             =   2640
            Width           =   2325
         End
         Begin VB.Label lblDob 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "DateofBirth"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   51
            Top             =   3120
            Width           =   1725
         End
         Begin VB.Label lblReligion 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Religion"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   50
            Top             =   3600
            Width           =   1365
         End
         Begin VB.Label lblCountry 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Country"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   49
            Top             =   4080
            Width           =   1365
         End
         Begin VB.Label lblState 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "StateofOrigin"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   48
            Top             =   4560
            Width           =   1365
         End
         Begin VB.Label lblLGA 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "LGA"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72720
            TabIndex        =   47
            Top             =   5040
            Width           =   1965
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   -74760
            TabIndex        =   46
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "L.G.A:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   -74760
            TabIndex        =   45
            Top             =   5040
            Width           =   765
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "State of Origin:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   -74760
            TabIndex        =   44
            Top             =   4560
            Width           =   1725
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   -74760
            TabIndex        =   43
            Top             =   4080
            Width           =   1005
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Religion:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   -74760
            TabIndex        =   42
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   -74760
            TabIndex        =   41
            Top             =   3120
            Width           =   1545
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Honours:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   -74760
            TabIndex        =   40
            Top             =   2640
            Width           =   1050
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   -74760
            TabIndex        =   39
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Titles:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   -74760
            TabIndex        =   38
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT DETAILS."
            BeginProperty Font 
               Name            =   "Angie Groovin"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   0
            Left            =   2400
            TabIndex        =   37
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblNofKinAd 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "NextofKinAdd"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   2880
            TabIndex        =   36
            Top             =   4080
            Width           =   4005
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Address"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   2880
            TabIndex        =   35
            Top             =   1200
            Width           =   4125
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPhone1 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Phone1"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   34
            Top             =   2160
            Width           =   2325
         End
         Begin VB.Label lblPhone2 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Phone2"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   33
            Top             =   2640
            Width           =   2325
         End
         Begin VB.Label lblNofKin 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "NextofKin"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   32
            Top             =   3120
            Width           =   4005
         End
         Begin VB.Label lblRelation 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "Relationship"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   31
            Top             =   3600
            Width           =   2325
         End
         Begin VB.Label lblNofKinPhone 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            DataField       =   "NextofKinPhone"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   30
            Top             =   5040
            Width           =   2325
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Phone No:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   28
            Top             =   5040
            Width           =   2580
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No1:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No2:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   2640
            Width           =   1350
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Name:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   25
            Top             =   3120
            Width           =   2115
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Address:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   24
            Top             =   4080
            Width           =   2355
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Address:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   1200
            Width           =   1950
         End
      End
      Begin VB.Label lbltag 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "H.O.D's PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   18
         Left            =   1560
         TabIndex        =   57
         Top             =   1080
         Width           =   6930
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   0
         Picture         =   "frmHODPI.frx":23505
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9975
      End
   End
   Begin VB.Frame fraEdit 
      BackColor       =   &H00FFFF00&
      Height          =   8655
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   9975
      Begin TabDlg.SSTab SSTab2 
         Height          =   6135
         Left            =   120
         TabIndex        =   61
         Top             =   1800
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "BIO-DATA"
         TabPicture(0)   =   "frmHODPI.frx":37B45
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtPassport"
         Tab(0).Control(1)=   "cmdNextPg"
         Tab(0).Control(2)=   "cmdUpLoad"
         Tab(0).Control(3)=   "txtLGA"
         Tab(0).Control(4)=   "txtState"
         Tab(0).Control(5)=   "txtCountry"
         Tab(0).Control(6)=   "txtReligion"
         Tab(0).Control(7)=   "txtDoB"
         Tab(0).Control(8)=   "txtHonours"
         Tab(0).Control(9)=   "txtSex"
         Tab(0).Control(10)=   "txtName"
         Tab(0).Control(11)=   "txtTitle"
         Tab(0).Control(12)=   "lbltag(28)"
         Tab(0).Control(13)=   "lbltag(27)"
         Tab(0).Control(14)=   "lbltag(26)"
         Tab(0).Control(15)=   "lbltag(25)"
         Tab(0).Control(16)=   "lbltag(24)"
         Tab(0).Control(17)=   "lbltag(23)"
         Tab(0).Control(18)=   "lbltag(22)"
         Tab(0).Control(19)=   "lbltag(21)"
         Tab(0).Control(20)=   "lbltag(20)"
         Tab(0).Control(21)=   "lbltag(19)"
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "CONTACT DETAILS"
         TabPicture(1)   =   "frmHODPI.frx":37B61
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lbltag(29)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbltag(30)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbltag(31)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lbltag(32)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbltag(33)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lbltag(34)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lbltag(35)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lbltag(36)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtAddress"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtPhone1"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txtPhone2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txtNofKin"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txtRelation"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtNofKinad"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtNofKinP"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         Begin VB.TextBox txtPassport 
            BackColor       =   &H00FFFFFF&
            DataField       =   "PassPort"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -68280
            TabIndex        =   81
            Top             =   3840
            Width           =   2895
         End
         Begin VB.CommandButton cmdNextPg 
            Caption         =   "GOTO NEXT PAGE >>>"
            Height          =   255
            Left            =   -71160
            TabIndex        =   10
            Top             =   5640
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpLoad 
            Caption         =   "Upload &PassPort >>>"
            Height          =   375
            Left            =   -67800
            TabIndex        =   9
            Top             =   4440
            Width           =   1935
         End
         Begin VB.TextBox txtNofKinP 
            DataField       =   "NextofKinPhone"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3000
            TabIndex        =   17
            Top             =   5040
            Width           =   2415
         End
         Begin VB.TextBox txtNofKinad 
            DataField       =   "NextofKinAdd"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   4080
            Width           =   3855
         End
         Begin VB.TextBox txtRelation 
            DataField       =   "Relationship"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3000
            TabIndex        =   15
            Top             =   3600
            Width           =   2415
         End
         Begin VB.TextBox txtNofKin 
            DataField       =   "NextofKin"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3000
            TabIndex        =   14
            Top             =   3120
            Width           =   3855
         End
         Begin VB.TextBox txtPhone2 
            DataField       =   "Phone2"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3000
            TabIndex        =   13
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtPhone1 
            DataField       =   "Phone1"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3000
            TabIndex        =   12
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtAddress 
            DataField       =   "Address"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txtLGA 
            DataField       =   "LGA"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   -72840
            TabIndex        =   8
            Top             =   5040
            Width           =   2055
         End
         Begin VB.TextBox txtState 
            DataField       =   "StateofOrigin"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   7
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox txtCountry 
            DataField       =   "Country"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   6
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox txtReligion 
            DataField       =   "Religion"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   5
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txtDoB 
            DataField       =   "DateofBirth"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   4
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox txtHonours 
            DataField       =   "Honours"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   3
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtSex 
            DataField       =   "Sex"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   2
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            DataField       =   "Names"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   1
            Top             =   1680
            Width           =   4695
         End
         Begin VB.TextBox txtTitle 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Title"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -72840
            TabIndex        =   0
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Address:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   36
            Left            =   360
            TabIndex        =   79
            Top             =   1200
            Width           =   1950
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Address:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   35
            Left            =   360
            TabIndex        =   78
            Top             =   4080
            Width           =   2355
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Name:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   34
            Left            =   360
            TabIndex        =   77
            Top             =   3120
            Width           =   2115
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No2:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   33
            Left            =   360
            TabIndex        =   76
            Top             =   2640
            Width           =   1350
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No1:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   32
            Left            =   360
            TabIndex        =   75
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next of Kin Phone No:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   31
            Left            =   360
            TabIndex        =   74
            Top             =   5040
            Width           =   2580
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   360
            TabIndex        =   73
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT DETAILS."
            BeginProperty Font 
               Name            =   "Angie Groovin"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   29
            Left            =   2520
            TabIndex        =   72
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Titles:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   -74760
            TabIndex        =   71
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   -74760
            TabIndex        =   70
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Honours:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   26
            Left            =   -74760
            TabIndex        =   69
            Top             =   2640
            Width           =   1050
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   -74760
            TabIndex        =   68
            Top             =   3120
            Width           =   1545
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Religion:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   24
            Left            =   -74760
            TabIndex        =   67
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   -74760
            TabIndex        =   66
            Top             =   4080
            Width           =   1005
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "State of Origin:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   -74760
            TabIndex        =   65
            Top             =   4560
            Width           =   1725
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "L.G.A:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   -74760
            TabIndex        =   64
            Top             =   5040
            Width           =   765
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   -74760
            TabIndex        =   63
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lbltag 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "BIO-DATA."
            BeginProperty Font 
               Name            =   "Angie Groovin"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   19
            Left            =   -72360
            TabIndex        =   62
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdBK 
         Caption         =   "&Back"
         Height          =   495
         Left            =   5160
         TabIndex        =   19
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   975
         Left            =   0
         Picture         =   "frmHODPI.frx":37B7D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9975
      End
      Begin VB.Label lbltag 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "H.O.D's PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Amelia"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   37
         Left            =   1560
         TabIndex        =   60
         Top             =   1080
         Width           =   6930
      End
   End
End
Attribute VB_Name = "frmHODPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProjectStaff.mdb;Persist Security Info=False"
End Function

Private Sub cmdBack_Click()
Me.Hide
frmHome.Show
frmHome.fraStaff.Visible = True
End Sub

Private Sub cmdBK_Click()
fraEdit.Visible = False
fraView.Visible = True
End Sub

Private Sub cmdEdit_Click()
cmdUpdate.Enabled = True
fraEdit.Visible = True
fraView.Visible = False
SSTab2.Tab = 0
End Sub

Private Sub cmdNextPg_Click()
SSTab2.Tab = 1
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo joe
GetConnect
Adodc1.Recordset.Update
cmdUpdate.Enabled = False
Exit Sub
joe:
MsgBox Err.Description
End Sub

Private Sub cmdUpLoad_Click()
cd1TesT.Flags = cdlOFNHideReadOnly
cd1TesT.Filter = "all files(*.*)|*.*|all pictures(*.pic)|*.pic|16 color bitmap(*.bmp)|*.bmp|jpeg(*.JPG;*.JPEG)|*.jgp;*.jgep|gif(*.gif)|*.gif|"
cd1TesT.ShowOpen
ImaPP.Picture = LoadPicture(cd1TesT.FileName)
txtPassport.Text = cd1TesT.FileName
End Sub

Private Sub Form_Load()
On Error GoTo kate
GetConnect
ImaPP.Picture = LoadPicture(txtPassport.Text)
Exit Sub
kate:
MsgBox Err.Description
End Sub
