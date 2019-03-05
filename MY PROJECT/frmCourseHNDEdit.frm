VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCourseHNDEdit 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12360
   Icon            =   "frmCourseHNDEdit.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   60
      Top             =   3600
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "HND I 1st Sem"
      TabPicture(0)   =   "frmCourseHNDEdit.frx":234CD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Adodc1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "List117"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List116"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "List115"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List114"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "List113"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List112"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "List111"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "List110"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "List109"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "List108"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "List107"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "List106"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "List105"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdOK"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "A"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "HND I 2nd Sem"
      TabPicture(1)   =   "frmCourseHNDEdit.frx":234E9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text8"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "List130"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "List129"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "List128"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "List127"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "List126"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "List125"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "List124"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "List123"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "List122"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "List121"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "List120"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "List119"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "List118"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdOK1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Adodc2"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "B"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "HND II 1st Sem"
      TabPicture(2)   =   "frmCourseHNDEdit.frx":23505
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text20"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text19"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text18"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Text17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text16"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text15"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "List143"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "List142"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "List141"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "List140"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "List139"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "List138"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "List137"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "List136"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "List135"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "List134"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "List133"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "List132"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "List131"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "cmdOK2"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Adodc3"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "C"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "HND II 2nd Sem"
      TabPicture(3)   =   "frmCourseHNDEdit.frx":23521
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text26"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text25"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Text24"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Text23"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Text22"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Text21"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "List1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "List2"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "List3"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "List4"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "List5"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "List6"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "List7"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "List8"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "List9"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "List10"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "List11"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "List12"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "List13"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "cmdOK3"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Adodc4"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "D"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).ControlCount=   22
      TabCaption(4)   =   "HND1(E) 1st S"
      TabPicture(4)   =   "frmCourseHNDEdit.frx":2353D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text27"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Text28"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Text29"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Text30"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Text31"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Text32"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdOK4"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "List26"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "List25"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "List24"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "List23"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "List22"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "List21"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "List20"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "List19"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "List18"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "List17"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "List16"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "List15"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "List14"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Adodc5"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Text33"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "Label1"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).ControlCount=   23
      TabCaption(5)   =   "HND1(E) 2nd S"
      TabPicture(5)   =   "frmCourseHNDEdit.frx":23559
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text34"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Text35"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Text36"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Text37"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Text38"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Text39"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "cmdOK5"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "List39"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "List38"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "List37"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "List36"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "List35"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "List34"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "List33"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "List32"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "List31"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "List30"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "List29"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "List28"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "List27"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Text40"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "Adodc6"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Label2"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).ControlCount=   23
      TabCaption(6)   =   "HND2(E) 1st S"
      TabPicture(6)   =   "frmCourseHNDEdit.frx":23575
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text41"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Text42"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Text43"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Text44"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Text45"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "cmdOK6"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "List52"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "List51"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "List50"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "List49"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "List48"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "List47"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "List46"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "List45"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "List44"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).Control(15)=   "List43"
      Tab(6).Control(15).Enabled=   0   'False
      Tab(6).Control(16)=   "List42"
      Tab(6).Control(16).Enabled=   0   'False
      Tab(6).Control(17)=   "List41"
      Tab(6).Control(17).Enabled=   0   'False
      Tab(6).Control(18)=   "List40"
      Tab(6).Control(18).Enabled=   0   'False
      Tab(6).Control(19)=   "Text46"
      Tab(6).Control(19).Enabled=   0   'False
      Tab(6).Control(20)=   "Adodc7"
      Tab(6).Control(20).Enabled=   0   'False
      Tab(6).Control(21)=   "Label3"
      Tab(6).Control(21).Enabled=   0   'False
      Tab(6).ControlCount=   22
      TabCaption(7)   =   "HND2(E) 2nd S"
      TabPicture(7)   =   "frmCourseHNDEdit.frx":23591
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Label4"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Adodc8"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Text50"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Text52"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "List53"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "List54"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "List55"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "List56"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "List57"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "List58"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "List59"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "List60"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).Control(12)=   "List61"
      Tab(7).Control(12).Enabled=   0   'False
      Tab(7).Control(13)=   "List62"
      Tab(7).Control(13).Enabled=   0   'False
      Tab(7).Control(14)=   "List63"
      Tab(7).Control(14).Enabled=   0   'False
      Tab(7).Control(15)=   "List64"
      Tab(7).Control(15).Enabled=   0   'False
      Tab(7).Control(16)=   "List65"
      Tab(7).Control(16).Enabled=   0   'False
      Tab(7).Control(17)=   "cmdOK7"
      Tab(7).Control(17).Enabled=   0   'False
      Tab(7).Control(18)=   "Text51"
      Tab(7).Control(18).Enabled=   0   'False
      Tab(7).Control(19)=   "Text49"
      Tab(7).Control(19).Enabled=   0   'False
      Tab(7).Control(20)=   "Text48"
      Tab(7).Control(20).Enabled=   0   'False
      Tab(7).Control(21)=   "Text47"
      Tab(7).Control(21).Enabled=   0   'False
      Tab(7).ControlCount=   22
      Begin VB.TextBox Text41 
         DataField       =   "STA411"
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
         TabIndex        =   51
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text42 
         DataField       =   "COM416"
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
         TabIndex        =   50
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text43 
         DataField       =   "COM415"
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
         TabIndex        =   49
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text44 
         DataField       =   "COM414"
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
         TabIndex        =   48
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text45 
         DataField       =   "COM413"
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
         TabIndex        =   47
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text47 
         DataField       =   "COM429"
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
         Left            =   8880
         TabIndex        =   58
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text48 
         DataField       =   "EED413"
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
         Left            =   8880
         TabIndex        =   57
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text49 
         DataField       =   "COM426"
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
         Left            =   8880
         TabIndex        =   56
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text51 
         DataField       =   "COM423"
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
         Left            =   8880
         TabIndex        =   54
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text34 
         DataField       =   "OTM320"
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
         TabIndex        =   44
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text35 
         DataField       =   "STA321"
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
         TabIndex        =   43
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text36 
         DataField       =   "COM326"
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
         TabIndex        =   42
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text37 
         DataField       =   "COM325"
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
         TabIndex        =   41
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text38 
         DataField       =   "COM323"
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
         TabIndex        =   40
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text39 
         DataField       =   "COM322"
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
         TabIndex        =   39
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text27 
         DataField       =   "OTM315"
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
         TabIndex        =   36
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text28 
         DataField       =   "sta314"
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
         TabIndex        =   35
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text29 
         DataField       =   "sta311"
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
         TabIndex        =   34
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text30 
         DataField       =   "com314"
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
         TabIndex        =   33
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text31 
         DataField       =   "com313"
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
         TabIndex        =   32
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text32 
         DataField       =   "com312"
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
         TabIndex        =   31
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK7 
         Caption         =   "&OK"
         Height          =   495
         Left            =   5400
         TabIndex        =   59
         Top             =   4800
         Width           =   1335
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
         ItemData        =   "frmCourseHNDEdit.frx":235AD
         Left            =   8880
         List            =   "frmCourseHNDEdit.frx":235B4
         TabIndex        =   172
         Top             =   960
         Width           =   2895
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":235C3
         Left            =   240
         List            =   "frmCourseHNDEdit.frx":235DC
         TabIndex        =   171
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":23612
         Left            =   1680
         List            =   "frmCourseHNDEdit.frx":2362B
         TabIndex        =   170
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":236EB
         Left            =   6720
         List            =   "frmCourseHNDEdit.frx":23704
         TabIndex        =   169
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":2371D
         Left            =   7080
         List            =   "frmCourseHNDEdit.frx":23736
         TabIndex        =   168
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":23750
         Left            =   7440
         List            =   "frmCourseHNDEdit.frx":2376C
         TabIndex        =   167
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":23788
         Left            =   8040
         List            =   "frmCourseHNDEdit.frx":237A1
         TabIndex        =   166
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":237C0
         Left            =   240
         List            =   "frmCourseHNDEdit.frx":237C7
         TabIndex        =   165
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":237D4
         Left            =   1680
         List            =   "frmCourseHNDEdit.frx":237DB
         TabIndex        =   164
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":237ED
         Left            =   6720
         List            =   "frmCourseHNDEdit.frx":237F4
         TabIndex        =   163
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":237FB
         Left            =   7080
         List            =   "frmCourseHNDEdit.frx":23802
         TabIndex        =   162
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":23809
         Left            =   7440
         List            =   "frmCourseHNDEdit.frx":23810
         TabIndex        =   161
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":23818
         Left            =   8040
         List            =   "frmCourseHNDEdit.frx":2381F
         TabIndex        =   160
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text52 
         DataField       =   "com422"
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
         Left            =   8880
         TabIndex        =   53
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text50 
         DataField       =   "COM424"
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
         Left            =   8880
         TabIndex        =   55
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK6 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   52
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
         ItemData        =   "frmCourseHNDEdit.frx":23829
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":23830
         TabIndex        =   158
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List51 
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
         ItemData        =   "frmCourseHNDEdit.frx":2383A
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":23841
         TabIndex        =   157
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List50 
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
         ItemData        =   "frmCourseHNDEdit.frx":23849
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":23850
         TabIndex        =   156
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List49 
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
         ItemData        =   "frmCourseHNDEdit.frx":23857
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":2385E
         TabIndex        =   155
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List48 
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
         ItemData        =   "frmCourseHNDEdit.frx":23865
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":2386C
         TabIndex        =   154
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List47 
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
         ItemData        =   "frmCourseHNDEdit.frx":2387E
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23885
         TabIndex        =   153
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
         Height          =   2910
         ItemData        =   "frmCourseHNDEdit.frx":23892
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":238A8
         TabIndex        =   152
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List45 
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
         ItemData        =   "frmCourseHNDEdit.frx":238C5
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":238DE
         TabIndex        =   151
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List44 
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
         ItemData        =   "frmCourseHNDEdit.frx":238F8
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":2390E
         TabIndex        =   150
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List43 
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
         ItemData        =   "frmCourseHNDEdit.frx":23924
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":2393A
         TabIndex        =   149
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List42 
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
         ItemData        =   "frmCourseHNDEdit.frx":23950
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23966
         TabIndex        =   148
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List41 
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
         ItemData        =   "frmCourseHNDEdit.frx":23A0A
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23A20
         TabIndex        =   147
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":23A54
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":23A5B
         TabIndex        =   146
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text46 
         DataField       =   "com412"
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
         TabIndex        =   46
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK5 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   45
         Top             =   4920
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
         ItemData        =   "frmCourseHNDEdit.frx":23A6A
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":23A71
         TabIndex        =   144
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List38 
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
         ItemData        =   "frmCourseHNDEdit.frx":23A7B
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":23A82
         TabIndex        =   143
         Top             =   960
         Width           =   615
      End
      Begin VB.ListBox List37 
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
         ItemData        =   "frmCourseHNDEdit.frx":23A8A
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":23A91
         TabIndex        =   142
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List36 
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
         ItemData        =   "frmCourseHNDEdit.frx":23A98
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":23A9F
         TabIndex        =   141
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox List35 
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
         ItemData        =   "frmCourseHNDEdit.frx":23AA6
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23AAD
         TabIndex        =   140
         Top             =   960
         Width           =   5055
      End
      Begin VB.ListBox List34 
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
         ItemData        =   "frmCourseHNDEdit.frx":23ABF
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23AC6
         TabIndex        =   139
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23AD3
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":23AEC
         TabIndex        =   138
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List32 
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
         ItemData        =   "frmCourseHNDEdit.frx":23B0E
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":23B2A
         TabIndex        =   137
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ListBox List31 
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
         ItemData        =   "frmCourseHNDEdit.frx":23B47
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":23B60
         TabIndex        =   136
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List30 
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
         ItemData        =   "frmCourseHNDEdit.frx":23B79
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":23B92
         TabIndex        =   135
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List29 
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
         ItemData        =   "frmCourseHNDEdit.frx":23BAB
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23BC4
         TabIndex        =   134
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ListBox List28 
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
         ItemData        =   "frmCourseHNDEdit.frx":23C78
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23C91
         TabIndex        =   133
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":23CCD
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":23CD4
         TabIndex        =   132
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text40 
         DataField       =   "com321"
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
         TabIndex        =   38
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   37
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
         ItemData        =   "frmCourseHNDEdit.frx":23CE3
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":23CEA
         TabIndex        =   130
         Top             =   960
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":23CF4
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":23CFB
         TabIndex        =   129
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":23D03
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":23D0A
         TabIndex        =   128
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":23D11
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":23D18
         TabIndex        =   127
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":23D1F
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23D26
         TabIndex        =   126
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":23D38
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23D3F
         TabIndex        =   125
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":23D4C
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":23D65
         TabIndex        =   124
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23D88
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":23DA4
         TabIndex        =   123
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23DC1
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":23DDA
         TabIndex        =   122
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23DF3
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":23E0C
         TabIndex        =   121
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23E25
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23E3E
         TabIndex        =   120
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         Height          =   3270
         ItemData        =   "frmCourseHNDEdit.frx":23EE6
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23EFF
         TabIndex        =   119
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":23F3B
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":23F42
         TabIndex        =   118
         Top             =   960
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adodc1 
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
      Begin VB.TextBox Text7 
         DataField       =   "OTM315"
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
         Left            =   -66120
         TabIndex        =   6
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "sta314"
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
         Left            =   -66120
         TabIndex        =   5
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         DataField       =   "sta311"
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
         Left            =   -66120
         TabIndex        =   4
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "com314"
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
         Left            =   -66120
         TabIndex        =   3
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "com313"
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
         Left            =   -66120
         TabIndex        =   2
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         DataField       =   "com312"
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
         Left            =   -66120
         TabIndex        =   1
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text26 
         DataField       =   "COM429"
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
         TabIndex        =   28
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text25 
         DataField       =   "EED413"
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
         TabIndex        =   27
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text24 
         DataField       =   "COM426"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text23 
         DataField       =   "COM424"
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
         TabIndex        =   25
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text22 
         DataField       =   "COM423"
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
         TabIndex        =   24
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text21 
         DataField       =   "COM422"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text20 
         DataField       =   "STA411"
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
         TabIndex        =   21
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text19 
         DataField       =   "COM416"
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
         TabIndex        =   20
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text18 
         DataField       =   "COM415"
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
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text17 
         DataField       =   "COM414"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text16 
         DataField       =   "COM413"
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
         TabIndex        =   17
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         DataField       =   "COM412"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         DataField       =   "OTM320"
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
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         DataField       =   "STA321"
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
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text12 
         DataField       =   "COM326"
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
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text11 
         DataField       =   "COM325"
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
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         DataField       =   "COM323"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         DataField       =   "COM322"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         DataField       =   "COM321"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "com311"
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
         Left            =   -66120
         TabIndex        =   0
         Top             =   1320
         Width           =   2895
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
         ItemData        =   "frmCourseHNDEdit.frx":23F51
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":23F58
         TabIndex        =   112
         Top             =   960
         Width           =   2895
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
         ItemData        =   "frmCourseHNDEdit.frx":23F67
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":23F80
         TabIndex        =   111
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":23FBC
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":23FD5
         TabIndex        =   110
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":2407D
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":24096
         TabIndex        =   109
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":240AF
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":240C8
         TabIndex        =   108
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":240E1
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":240FD
         TabIndex        =   107
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":2411A
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":24133
         TabIndex        =   106
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":24156
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":2415D
         TabIndex        =   105
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":2416A
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":24171
         TabIndex        =   104
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":24183
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":2418A
         TabIndex        =   103
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":24191
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":24198
         TabIndex        =   102
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":2419F
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":241A6
         TabIndex        =   101
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":241AE
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":241B5
         TabIndex        =   100
         Top             =   960
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":241BF
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":241C6
         TabIndex        =   99
         Top             =   960
         Width           =   2895
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
         ItemData        =   "frmCourseHNDEdit.frx":241D5
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":241EE
         TabIndex        =   98
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":2422A
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":24243
         TabIndex        =   97
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":242F7
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":24310
         TabIndex        =   96
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":24329
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":24342
         TabIndex        =   95
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":2435B
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":24377
         TabIndex        =   94
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":24394
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":243AD
         TabIndex        =   93
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":243CF
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":243D6
         TabIndex        =   92
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":243E3
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":243EA
         TabIndex        =   91
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":243FC
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":24403
         TabIndex        =   90
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":2440A
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":24411
         TabIndex        =   89
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":24418
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":2441F
         TabIndex        =   88
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":24427
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":2442E
         TabIndex        =   87
         Top             =   960
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":24438
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":2443F
         TabIndex        =   86
         Top             =   960
         Width           =   2895
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
         ItemData        =   "frmCourseHNDEdit.frx":2444E
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":24464
         TabIndex        =   85
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":24498
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":244AE
         TabIndex        =   84
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":24552
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":24568
         TabIndex        =   83
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":2457E
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":24594
         TabIndex        =   82
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":245AA
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":245C3
         TabIndex        =   81
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":245DD
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":245F3
         TabIndex        =   80
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":24610
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":24617
         TabIndex        =   79
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":24624
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":2462B
         TabIndex        =   78
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":2463D
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":24644
         TabIndex        =   77
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":2464B
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":24652
         TabIndex        =   76
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":24659
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":24660
         TabIndex        =   75
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":24668
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":2466F
         TabIndex        =   74
         Top             =   960
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":24679
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":24680
         TabIndex        =   73
         Top             =   960
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":2468A
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":24691
         TabIndex        =   72
         Top             =   960
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":24699
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":246A0
         TabIndex        =   71
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":246A7
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":246AE
         TabIndex        =   70
         Top             =   960
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":246B5
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":246BC
         TabIndex        =   69
         Top             =   960
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":246CE
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":246D5
         TabIndex        =   68
         Top             =   960
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":246E2
         Left            =   -66960
         List            =   "frmCourseHNDEdit.frx":246FB
         TabIndex        =   67
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   855
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
         ItemData        =   "frmCourseHNDEdit.frx":2471A
         Left            =   -67560
         List            =   "frmCourseHNDEdit.frx":24736
         TabIndex        =   66
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   615
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
         ItemData        =   "frmCourseHNDEdit.frx":24752
         Left            =   -67920
         List            =   "frmCourseHNDEdit.frx":2476B
         TabIndex        =   65
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":24785
         Left            =   -68280
         List            =   "frmCourseHNDEdit.frx":2479E
         TabIndex        =   64
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   375
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
         ItemData        =   "frmCourseHNDEdit.frx":247B7
         Left            =   -73320
         List            =   "frmCourseHNDEdit.frx":247D0
         TabIndex        =   63
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   5055
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
         ItemData        =   "frmCourseHNDEdit.frx":24890
         Left            =   -74760
         List            =   "frmCourseHNDEdit.frx":248A9
         TabIndex        =   62
         Tag             =   "LIST OF COURSES"
         Top             =   1320
         Width           =   1455
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
         ItemData        =   "frmCourseHNDEdit.frx":248DF
         Left            =   -66120
         List            =   "frmCourseHNDEdit.frx":248E6
         TabIndex        =   61
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   7
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK1 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   15
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK2 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   22
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK3 
         Caption         =   "&OK"
         Height          =   495
         Left            =   -69600
         TabIndex        =   29
         Top             =   4800
         Width           =   1335
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
      Begin VB.TextBox Text33 
         DataField       =   "com311"
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
         TabIndex        =   30
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label4 
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
         Left            =   3360
         TabIndex        =   173
         Top             =   480
         Width           =   6000
      End
      Begin VB.Label Label3 
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
         Left            =   -71760
         TabIndex        =   159
         Top             =   480
         Width           =   5805
      End
      Begin VB.Label Label2 
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
         TabIndex        =   145
         Top             =   480
         Width           =   6000
      End
      Begin VB.Label Label1 
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
         Left            =   -71760
         TabIndex        =   131
         Top             =   480
         Width           =   5700
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
         TabIndex        =   116
         Top             =   480
         Width           =   4005
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
         TabIndex        =   115
         Top             =   480
         Width           =   4110
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
         Left            =   -71520
         TabIndex        =   114
         Top             =   480
         Width           =   4305
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
         TabIndex        =   113
         Top             =   480
         Width           =   4305
      End
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmCourseHNDEdit.frx":248F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
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
      Left            =   3120
      TabIndex        =   117
      Top             =   2880
      Width           =   6240
   End
End
Attribute VB_Name = "frmCourseHNDEdit"
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
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 0
End Sub

Private Sub cmdOK1_Click()
GetConnect2
Adodc2.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 1
End Sub

Private Sub cmdOK2_Click()
GetConnect3
Adodc3.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 2
End Sub

Private Sub cmdOK3_Click()
GetConnect4
Adodc4.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 3
End Sub

Private Sub cmdOK4_Click()
GetConnect5
Adodc5.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 4
End Sub

Private Sub cmdOK5_Click()
GetConnect6
Adodc6.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 5
End Sub

Private Sub cmdOK6_Click()
GetConnect7
Adodc7.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 6
End Sub

Private Sub cmdOK7_Click()
GetConnect8
Adodc8.Recordset.Update
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 7
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
