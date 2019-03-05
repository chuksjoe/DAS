VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChangePass 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Passworrd"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2280
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "tblPassword"
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtCPassW 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtNpassW 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtVpassW 
      DataField       =   "password"
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE YOUR PASSWORD."
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
      Left            =   840
      TabIndex        =   8
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frmChangePass.frx":234CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password:"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password:"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   2025
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetConnect()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:" & App.Path & "\ProLecturers.mdb;Persist Security Info=False"
End Function
Private Sub cmdBack_Click()
Me.Hide
frmHome.Show
End Sub

Private Sub cmdChange_Click()
GetConnect
If txtCPassW.Text = frmLogin.lblpassword.Caption Then
If txtNpassW.Text = txtVpassW.Text Then
frmLogin.lblpassword.Caption = txtNpassW.Text
Adodc1.Recordset.Update
mes = MsgBox("Password Changed Successfully", vbOKOnly, "CONFIRMATION")
txtCPassW.Text = ""
txtNpassW.Text = ""
txtVpassW.Text = ""
txtCPassW.Enabled = False
txtNpassW.Enabled = False
txtVpassW.Enabled = False
Else
mes2 = MsgBox("Your new password does not match", vbOKOnly + vbExclamation, "WARNING")
txtNpassW.Text = ""
txtVpassW.Text = ""
txtNpassW.SetFocus
End If
Else
mes3 = MsgBox("Incorrect password, try again.", vbOKOnly + vbExclamation, "WARNING")
txtCPassW.Text = ""
txtNpassW.Text = ""
txtVpassW.Text = ""
txtCPassW.SetFocus
End If
End Sub

