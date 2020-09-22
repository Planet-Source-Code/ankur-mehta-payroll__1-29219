VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Login into our system"
   ClientHeight    =   2595
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7830
   Icon            =   "payrolllogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1533.212
   ScaleMode       =   0  'User
   ScaleWidth      =   7351.946
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton Command2 
         Caption         =   "Return"
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   1560
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label2 
         Caption         =   "Just enter your name and you password in the Diagrams shown below "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Please Enter your name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2400
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Please enter your password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Line Line1 
         X1              =   4920
         X2              =   6000
         Y1              =   1080
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Your name"
         Height          =   255
         Left            =   6120
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   4920
         X2              =   5880
         Y1              =   1800
         Y2              =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "Your password"
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Text            =   "manager"
      Top             =   720
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   6720
      TabIndex        =   4
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   6720
      TabIndex        =   5
      Top             =   720
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "manager"
      Top             =   1320
      Width           =   2325
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37046
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Today's date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to professional payroll software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Please Enter your name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Please enter your password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   2520
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
End 'exits program
End Sub

Private Sub cmdOK_Click()
Dim aa As String, pp As String
aa = LCase(txtUserName(0).Text) 'sets ualue of aa
pp = LCase(txtPassword(0).Text) 'sets ualue of pp
If aa = "manager" And pp = "manager" Then 'does boolean operators
   Form1.Show 'shows the form
   frmLogin.Hide 'hides the form
   
End If 'ends the end if procedure

End Sub

Private Sub Command1_Click()
Frame1.Visible = True 'shows the frame

End Sub

Private Sub Command2_Click()
Frame1.Visible = False 'hides the frame

End Sub

