VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8790
   Icon            =   "shoeemployees.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   615
         Left            =   7560
         TabIndex        =   58
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   615
         Left            =   6720
         TabIndex        =   57
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   615
         Left            =   5880
         TabIndex        =   56
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   615
         Left            =   5040
         TabIndex        =   55
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   615
         Left            =   4200
         TabIndex        =   54
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   615
         Left            =   3360
         TabIndex        =   53
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   615
         Left            =   2520
         TabIndex        =   52
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   1680
         TabIndex        =   51
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Left            =   840
         TabIndex        =   50
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   0
         TabIndex        =   49
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker text48 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37046
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total allowance"
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
      Index           =   15
      Left            =   4320
      TabIndex        =   47
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Taxable income"
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
      Index           =   12
      Left            =   4320
      TabIndex        =   46
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   11
      Left            =   4320
      TabIndex        =   45
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   10
      Left            =   4320
      TabIndex        =   44
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   9
      Left            =   4320
      TabIndex        =   43
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total tax"
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
      Index           =   8
      Left            =   4320
      TabIndex        =   42
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   13
      Left            =   4320
      TabIndex        =   41
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   14
      Left            =   4320
      TabIndex        =   40
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   21
      Left            =   4320
      TabIndex        =   39
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   22
      Left            =   4320
      TabIndex        =   38
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   23
      Left            =   4320
      TabIndex        =   37
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Telephone"
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
      Index           =   24
      Left            =   4320
      TabIndex        =   36
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Address "
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
      Index           =   25
      Left            =   120
      TabIndex        =   35
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   26
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Married "
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
      Index           =   27
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Widow"
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
      Index           =   28
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gender"
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
      Index           =   29
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   30
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Age"
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
      Index           =   31
      Left            =   120
      TabIndex        =   27
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Disable "
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
      Index           =   32
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Index           =   33
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   24
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   23
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   22
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6120
      TabIndex        =   21
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   20
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   6120
      TabIndex        =   19
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   6120
      TabIndex        =   18
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   6120
      TabIndex        =   17
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   6120
      TabIndex        =   16
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   6240
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   2040
      TabIndex        =   14
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Job Title"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Menu ss 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu dfd 
         Caption         =   "Menu"
      End
      Begin VB.Menu dfcd 
         Caption         =   "Save"
      End
      Begin VB.Menu gp 
         Caption         =   "Generate Payslip"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text = "W" Then 'sets the boolean variable
  Label3(10) = "£8.00" 'sets the label to 8 pound
ElseIf Combo1.Text = "S" Then 'sets the boolean variable
Label3(10) = "£6.00" 'sets the label to 8 pound
ElseIf Combo1.Text = "U" Then 'sets the boolean variable
Label3(10) = "£4" 'sets the label to 8 pound
End If 'ends the end if statement
End Sub

Private Sub dfcd_Click()
On Error GoTo errorhandler 'handles any sort of error
Dim X As Integer 'declares the variable


aa = FreeFile 'sets aa variable to freefile
                                        id(save) = Form5.Text1(0) 'set the value of the variable from text box
                                        fname(save) = Form5.Text2(0) 'set the value of the variable from text box
                                        sname(save) = Form5.Text3(0) 'set the value of the variable from text box
                                        jtitle(save) = Form5.Combo1 'set the value of the variable from combo box
                                        disable(save) = Form5.Combo5 'set the value of the variable from combo box
                                        age(save) = Form5.Text5(0) 'set the value of the variable from text box
                                        payrate(save) = Form5.Label3(10) 'set the value of the variable from label box
                                        gender(save) = Form5.Combo2 'set the value of the variable from label box
                                        widow(save) = Form5.Combo3 'set the value of the variable from label box
                                        married(save) = Form5.Combo4 'set the value of the variable from label box
                                        ninumber(save) = Form5.Text10(0) 'set the value of the variable from text box
                                        address(save) = Form5.Text1(1) 'set the value of the variable from text box
                                        telephone(save) = Form5.Text2(1) 'set the value of the variable from text box
                                        cpaydate(save) = Form5.Label3(9) 'set the value of the variable from label box
                                        lpaydate(save) = Form5.text48 'set the value of the variable from text box
                                        Totalpay(save) = Form5.Label3(6) 'set the value of the variable from label box
                                        allowance(save) = Form5.Label3(5) 'set the value of the variable from label box
                                        taxableincome(save) = Form5.Label3(7) 'set the value of the variable from label box
                                        Nicontribution(save) = Form5.Label3(8) 'set the value of the variable from label box
                                        tax(save) = Form5.Label3(0) 'set the value of the variable from label box
                                        totalnicontribution(save) = Form5.Label3(1) 'set the value of the variable from label box
                                        totaltax(save) = Form5.Label3(2) 'set the value of the variable from label box
                                        totalearning(save) = Form5.Label3(3) 'set the value of the variable from label box
                                        netpay(save) = Form5.Label3(4) 'set the value of the variable from label box
                                        
Open App.Path + "\payroll.dLL" For Output As aa 'opens the file
For X = 1 To 1 'starts x variable loop
 
 Write #aa, id(X), fname(X), sname(X), jtitle(X), disable(X), age(X), payrate(X), gender(X), widow(X), married(X), ninumber(X), address(X), telephone(X), cpaydate(X), lpaydate(X), Totalpay(X), allowance(X), taxableincome(X), Nicontribution(X), tax(X), totalnicontribution(X), totaltax(X), totalearning(X), netpay(X) 'writes to the file

Next X 'nest loop
Close #aa 'closes file
errorhandler:
        Debug.Print "error" 'prints message in debug
        Close #aa 'closes file
End Sub

Private Sub dfd_Click()
Form1.Show 'shows form
Form5.Hide 'hides form

End Sub

Private Sub Form_Load()





Combo1.AddItem "S" 'adds the item in the combo box
Combo1.AddItem "U" 'adds the item in the combo box
Combo1.AddItem "W" 'adds the item in the combo box
Combo5.AddItem "Yes" 'adds the item in the combo box
Combo5.AddItem "No" 'adds the item in the combo box
Combo2.AddItem "M" 'adds the item in the combo box
Combo2.AddItem "F" 'adds the item in the combo box
Combo4.AddItem "Yes" 'adds the item in the combo box
Combo4.AddItem "No" 'adds the item in the combo box
Combo3.AddItem "Yes" 'adds the item in the combo box
Combo3.AddItem "No" 'adds the item in the combo box






End Sub

Private Sub gp_Click()
Form7.Text3(6) = Text3(0).Text 'set the value of the text box
Form7.Text1(3) = Text2(0).Text 'set the value of the text box
Form7.Text2(1) = Text1(0).Text 'set the value of the text box
Form7.Text1(1) = text48 'set the value of the text box
Form7.Text2(0) = Label3(9) 'set the value of the text box
Form7.Text3(16) = Text10(0).Text 'set the value of the text box'set the value of the text box
Form7.Text1(6) = Label3(1) 'set the value of the text box
Form7.Text1(2) = Label3(3) 'set the value of the text box
Form7.Text2(3) = Label3(10) 'set the value of the text box
Form7.Text3(11) = Label3(0) 'set the value of the text box
Form7.Text1(4) = Label3(6) 'set the value of the text box
Form7.Text3(9) = Label3(4) 'set the value of the text box
Form7.Show 'shows the form
Form7.Command1.SetFocus 'focus on the particular command button



End Sub

Private Sub showbydetails_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
