VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000009&
   Caption         =   "Payslip"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   Picture         =   "payslip7.frx":0000
   ScaleHeight     =   7770
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Pay Rate"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Total pay"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   13
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Current Pay date "
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Employee ID"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   11
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Name"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Surname"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Last Pay date "
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Current Pay date "
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Total pay"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "NI number"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Total NI contibution"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "NI contribution"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   4800
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "Total earning"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   16
      Left            =   2280
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   7080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pay rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   30
      Left            =   480
      TabIndex        =   30
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   21
      Left            =   6600
      TabIndex        =   29
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last pay date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   22
      Left            =   1200
      TabIndex        =   28
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current pay date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   23
      Left            =   5880
      TabIndex        =   27
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee ID "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total NI contribution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   24
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tax "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   4440
      TabIndex        =   23
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NI contribution "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   2160
      TabIndex        =   22
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total earning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   6000
      TabIndex        =   21
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Net pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   8400
      TabIndex        =   20
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NI number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   26
      Left            =   6000
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Surname "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   33
      Left            =   360
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "270 Lady Margaret Road"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Southall         UD1 245"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Middlesex London UK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False 'makes command visiblility false
Command2.Visible = False 'makes command visiblility false
Form7.PrintForm 'prints the form
Command1.Visible = True 'makes command visiblility true
Command2.Visible = True 'makes command visiblility true
End Sub

Private Sub Command2_Click()
Unload Me 'hides the form

End Sub

