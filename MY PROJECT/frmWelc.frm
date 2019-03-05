VERSION 5.00
Begin VB.Form frmWelc 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   Icon            =   "frmWelc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHint 
      Caption         =   "&Hint"
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
      Left            =   7080
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      DownPicture     =   "frmWelc.frx":234CD
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
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderWidth     =   12
      X1              =   120
      X2              =   10560
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderWidth     =   12
      X1              =   120
      X2              =   10560
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   0
      Picture         =   "frmWelc.frx":331DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTAL ADMINISTRATIVE SYSTEM"
      BeginProperty Font 
         Name            =   "Amelia"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   10455
   End
End
Attribute VB_Name = "frmWelc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
msg = MsgBox("Do you want to Exit?", vbYesNo + vbQuestion, "EXIT")
If msg = vbYes Then
End
End If
End Sub

Private Sub cmdHint_Click()
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

Private Sub cmdLogin_Click()
frmLogin.Show
End Sub
