VERSION 5.00
Begin VB.Form frmCourses 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courses"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   Icon            =   "frmCourses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHNDCourses 
      Caption         =   "Higher National Diploma"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton cmdNDCourses 
      Caption         =   "National Diploma"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "frmCourses.frx":234CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTAL COURSES"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   4365
   End
End
Attribute VB_Name = "frmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Me.Hide
frmHome.Show
End Sub

Private Sub cmdHNDCourses_Click()
Me.Hide
frmCourseHND.Show
frmCourseHND.SSTab1.Tab = 0
End Sub

Private Sub cmdNDCourses_Click()
Me.Hide
frmCourseND.Show
frmCourseND.SSTab1.Tab = 0
End Sub
