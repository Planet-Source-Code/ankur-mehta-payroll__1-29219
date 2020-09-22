VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "setup1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   3480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   1680
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Max             =   2500
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Please wait program installing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   1560
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select the application path"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   6495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Please select the application path and click next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   3855
      Left            =   -1200
      TabIndex        =   1
      Top             =   1800
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Quit"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Click on next to button to install the software "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome to Payroll software Version 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True

End Sub

Private Sub Command3_Click()
Form2.Show

End Sub

Private Sub Command4_Click()
Frame2.Visible = False
Frame3.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()


Label1.Top = Me.Top + 200
Frame1.Top = (Me.Height / 2) - 1500
Frame1.Left = (Me.Width / 2) - 3600
Frame2.Top = (Me.Height / 2) - 1500
Frame2.Left = (Me.Width / 2) - 3600
Frame3.Top = (Me.Height / 2) - 1500
Frame3.Left = (Me.Width / 2) - 3600
End Sub

Private Sub Timer1_Timer()
On Error GoTo errorhandler
Dim X As Integer
Dim fsys As Object
Set fsys = CreateObject("scripting.filesystemobject")
path = path + "\"


fsys.copyfolder "a:\Payroll software", path








For X = 1 To 2500
 ProgressBar1.Value = ProgressBar1.Value + 1
Next X

MsgBox "Program installed at " + path + "Payroll software\Payroll.exe", vbInformation, "Path"
End
errorhandler:
    ProgressBar1.Enabled = False
    
    
End Sub
