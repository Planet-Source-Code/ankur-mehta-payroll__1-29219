VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tax rates"
   ClientHeight    =   4800
   ClientLeft      =   1395
   ClientTop       =   1965
   ClientWidth     =   8610
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   4800
   ScaleWidth      =   8610
   Begin VB.Frame Frame4 
      Caption         =   "Allowances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4440
      TabIndex        =   23
      Top             =   2760
      Width           =   3975
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Text            =   "Text14"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Text            =   "Text13"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Text            =   "Text12"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Text            =   "Text11"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Disable allowance"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Widow allowance"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Age allowance"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Married allowance"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Company optional deductions "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   4335
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Text            =   "Text10"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Text            =   "Text9"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Compant union rate "
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Company pension rate"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "NI deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Width           =   4095
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Text            =   "Text8"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text7"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Greater than 500"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Greater than 114"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Greater than 70"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Less than 70"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tax rates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Greater than 500"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Greater than 114"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Greater than 70"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Less than 70"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Menu s 
      Caption         =   "File"
      Begin VB.Menu w 
         Caption         =   "Save"
      End
      Begin VB.Menu q 
         Caption         =   "Main menu"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub q_Click()
Form4.Hide 'hides the form
Form1.Show 'shows the form

End Sub

Private Sub w_Click()
Dim aa As Integer 'declares the variable
aa = FreeFile 'sets aa variable to free file
Open App.Path + "\taxrates.dLL" For Output As #aa 'opens the file
Print #aa, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text 'puts the values to the particular text boxes
Close #aa 'closes the file

End Sub
