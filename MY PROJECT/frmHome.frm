VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6345
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHome 
      BackColor       =   &H0080FF80&
      Height          =   6735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdStuRecord 
         Caption         =   "Students Record"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   0
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton cmdStaffRecord 
         Caption         =   "Staff Record"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   1
         Top             =   3000
         Width           =   4095
      End
      Begin VB.CommandButton cmdCourses 
         Caption         =   "Courses"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   2
         Top             =   3960
         Width           =   4095
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangePW 
         Caption         =   "Change Admin Password"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   3
         Top             =   4920
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   0
         Picture         =   "frmHome.frx":234CD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.Frame fraSTUDENT 
      BackColor       =   &H0080FFFF&
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdBk2Home 
         Caption         =   "&Back"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdNd 
         Caption         =   "ND"
         BeginProperty Font 
            Name            =   "Cosmic"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   840
         TabIndex        =   14
         Top             =   3120
         Width           =   4695
      End
      Begin VB.CommandButton cmdhnd 
         Caption         =   "HND"
         BeginProperty Font 
            Name            =   "Cosmic"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   840
         TabIndex        =   13
         Top             =   840
         Width           =   4695
      End
      Begin VB.Image Image3 
         Height          =   6660
         Left            =   0
         Picture         =   "frmHome.frx":4729B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6405
      End
   End
   Begin VB.Frame fraStaff 
      BackColor       =   &H0000FF00&
      Height          =   6735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdHod 
         Caption         =   "HOD"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   11
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton cmdAcdSatff 
         Caption         =   "Academic Staff"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   10
         Top             =   3000
         Width           =   4095
      End
      Begin VB.CommandButton cmdNonAcdStaff 
         Caption         =   "Non-Academic Staff"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   9
         Top             =   3960
         Width           =   4095
      End
      Begin VB.CommandButton cmdBk2Home 
         Caption         =   "&Back"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdStaffTable 
         Caption         =   "Staff Record Table"
         BeginProperty Font 
            Name            =   "KabarettD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   7
         Top             =   4920
         Width           =   4095
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   0
         Picture         =   "frmHome.frx":83295
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6360
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuStuRecord 
      Caption         =   "Student Record"
      Begin VB.Menu mnuHND 
         Caption         =   "HND"
         Begin VB.Menu mnuHNDM 
            Caption         =   "Morning Student"
         End
         Begin VB.Menu mnuHNDE 
            Caption         =   "Evening Student"
         End
      End
      Begin VB.Menu mnuND 
         Caption         =   "ND"
         Begin VB.Menu mnuNDM 
            Caption         =   "Morning Student"
         End
         Begin VB.Menu mnuNDE 
            Caption         =   "Evening Student"
         End
      End
   End
   Begin VB.Menu mnuStaffRecord 
      Caption         =   "Staff Record"
      Begin VB.Menu mnuHOD 
         Caption         =   "HOD"
      End
      Begin VB.Menu mnuAStaff 
         Caption         =   "Academic Staff"
      End
      Begin VB.Menu mnuNAStaff 
         Caption         =   "Non-Academic Staff"
      End
   End
   Begin VB.Menu mnuCourses 
      Caption         =   "Courses"
      Begin VB.Menu mnuHNDCourses 
         Caption         =   "HND Courses"
      End
      Begin VB.Menu mnuNDCourse 
         Caption         =   "ND Courses"
      End
   End
   Begin VB.Menu mnuChangePW 
      Caption         =   "Change Password"
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcdSatff_Click()
Me.Hide
frmAcadStaffPI.Show
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmWelc.Show
End Sub

Private Sub cmdBk2Home_Click(Index As Integer)
fraStaff.Visible = False
fraSTUDENT.Visible = False
fraHome.Visible = True
End Sub

Private Sub cmdChangePW_Click()
Me.Hide
frmChangePass.Show
frmChangePass.txtCPassW.Enabled = True
frmChangePass.txtCPassW.SetFocus
frmChangePass.txtVpassW = ""
End Sub

Private Sub cmdCourses_Click()
Me.Hide
frmCourses.Show
End Sub

Private Sub cmdhnd_Click()
Me.Hide
frmHND.Show
End Sub

Private Sub cmdHod_Click()
Me.Hide
frmHODPI.Show
frmHODPI.SSTab1.Tab = 0
End Sub

Private Sub cmdNd_Click()
Me.Hide
frmNd.Show
End Sub

Private Sub cmdNonAcdStaff_Click()
Me.Hide
frmNonAcadStaffPI.Show
End Sub

Private Sub cmdStaffRecord_Click()
fraHome.Visible = False
fraStaff.Visible = True
End Sub

Private Sub cmdStaffTable_Click()
Me.Hide
frmStaffTable.Show
End Sub

Private Sub cmdStuRecord_Click()
fraHome.Visible = False
fraSTUDENT.Visible = True
End Sub

Private Sub mnuAbout_Click()
msg = ":::::::::::::::::::::::::::::::::::::::::::::INTRO::::::::::::::::::::::::::::::::::::::::::::::::::" + vbCr
msg = msg + "This is Orjiakor Chukwunonso's Project." + vbCr + "It is a DEPARTMENTAL ADMINISTRATTIVE SYSTEM SOFTWARE" + vbCr
msg = msg + "used for managing student records, that is their personal info" + vbCr + "(i.e, their Bio-Data, contact Info and their parent contact Info)" + vbCr
msg = msg + "and their Result Details and it is used for all the students" + vbCr + " in the Department." + vbCr
msg = msg + "Other features that are included in this my Software are:" + vbCr
msg = msg + "   -Staff(Academic and non-Academic) Bio_Data and Contact info" + vbCr
msg = msg + "   -Courses Studied in all levels in the Department and the Lecturers" + vbCr
msg = msg + "   -and Resetable Password for accessing the Software" + vbCr
msg = msg + "Some other features will be integrated into the newer version."
MsgBox msg, vbOKOnly, "Hint"
End Sub

Private Sub mnuAStaff_Click()
Me.Hide
frmAcadStaffPI.Show
End Sub

Private Sub mnuChangePW_Click()
Me.Hide
frmChangePass.Show
frmChangePass.txtVpassW.Text = ""
End Sub

Private Sub mnuExit_Click()
msg = MsgBox("Do you want to Exit?", vbYesNo + vbQuestion, "EXIT")
If msg = vbYes Then
End
End If
End Sub

Private Sub mnuHNDCourses_Click()
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 0
End Sub

Private Sub mnuHNDE_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDE.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab2.Tab = 0
End Sub

Private Sub mnuHNDM_Click()
Me.Hide
frmHND.Show
frmHND.fraHNDM.Visible = True
frmHND.fraHND.Visible = False
frmHND.SSTab1.Tab = 0
End Sub

Private Sub mnuHOD_Click()
Me.Hide
frmHODPI.Show
frmHODPI.SSTab1.Tab = 0
End Sub

Private Sub mnuNAStaff_Click()
Me.Hide
frmNonAcadStaffPI.Show
End Sub

Private Sub mnuNDCourse_Click()
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 0
End Sub

Private Sub mnuNDE_Click()
Me.Hide
frmNd.Show
frmNd.fraND.Visible = False
frmNd.fraNDE.Visible = True
frmNd.SSTab1.Tab = 0
End Sub

Private Sub mnuNDM_Click()
Me.Hide
frmNd.Show
frmNd.fraND.Visible = False
frmNd.fraNDM.Visible = True
frmNd.SSTab2.Tab = 0
End Sub
