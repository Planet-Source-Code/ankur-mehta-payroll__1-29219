VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Employees details"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11715
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   11715
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1800
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1800
         TabIndex        =   35
         Top             =   5760
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   5160
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   32
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   31
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   30
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   29
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   28
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   25
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "Pension"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   3360
         Width           =   4215
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "Union"
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
         Left            =   3480
         TabIndex        =   22
         Top             =   3960
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show deductions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4560
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hide deductions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   20
         Top             =   5520
         Visible         =   0   'False
         Width           =   3975
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
         Height          =   375
         Index           =   22
         Left            =   3480
         TabIndex        =   53
         Top             =   2760
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
         Height          =   375
         Index           =   23
         Left            =   3480
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   3480
         TabIndex        =   51
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   3480
         TabIndex        =   50
         Top             =   840
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
         Height          =   375
         Index           =   0
         Left            =   -120
         TabIndex        =   49
         Top             =   120
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
         Height          =   375
         Index           =   2
         Left            =   -120
         TabIndex        =   48
         Top             =   840
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
         Height          =   375
         Index           =   26
         Left            =   3480
         TabIndex        =   47
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   -120
         TabIndex        =   46
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   -120
         TabIndex        =   45
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   -120
         TabIndex        =   44
         Top             =   4560
         Width           =   1695
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
         Height          =   375
         Index           =   30
         Left            =   -120
         TabIndex        =   43
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   -120
         TabIndex        =   42
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   -120
         TabIndex        =   41
         Top             =   2760
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
         Height          =   375
         Index           =   33
         Left            =   -120
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   39
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   1800
         TabIndex        =   38
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   0
         TabIndex        =   37
         Top             =   2040
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   62
         Top             =   5520
         Width           =   135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   61
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   60
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   59
         Top             =   4920
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   58
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   57
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   56
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   55
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "£"
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   17
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
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1695
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
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   1200
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
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   13
         Top             =   2520
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
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   12
         Top             =   3120
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
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gross pay"
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
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2040
         TabIndex        =   5
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   4
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   3
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   2
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu ww 
         Caption         =   "Main Menu"
         Index           =   2
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id 'declares variable  which will help the program to reterive information from files


Dim fname 'declares variable  which will help the program to reterive information from files


Dim sname 'declares variable  which will help the program to reterive information from files
Dim jtitle 'declares variable  which will help the program to reterive information from files


Dim disable 'declares variable  which will help the program to reterive information from files
Dim age 'declares variable  which will help the program to reterive information from files

Dim payrate 'declares variable  which will help the program to reterive information from files


Dim gender 'declares variable  which will help the program to reterive information from files

Dim widow 'declares variable  which will help the program to reterive information from files
Dim married 'declares variable  which will help the program to reterive information from files
Dim ninumber 'declares variable  which will help the program to reterive information from files

Dim address 'declares variable  which will help the program to reterive information from files

Dim telephone 'declares variable  which will help the program to reterive information from files

Dim cpaydate 'declares variable  which will help the program to reterive information from files
Dim lpaydate 'declares variable  which will help the program to reterive information from files
Dim Totalpay 'declares variable  which will help the program to reterive information from files
Dim allowance 'declares variable  which will help the program to reterive information from files
Dim taxableincome 'declares variable  which will help the program to reterive information from files
Dim Nicontribution 'declares variable  which will help the program to reterive information from files
Dim tax 'declares variable  which will help the program to reterive information from files
Dim totalnicontribution 'declares variable  which will help the program to reterive information from files
Dim totaltax 'declares variable  which will help the program to reterive information from files
Dim totalearning 'declares variable  which will help the program to reterive information from files
Dim netpay 'declares variable  which will help the program to reterive information from files


Private Sub Command1_Click()
Dim aa As Integer 'declares variable called aa to open the file
Dim DEDUCTIONS As Double 'declares variable to count deduction
aa = FreeFile
Dim a As Double 'declares variable
Dim b As Double 'declares variable
Dim c As Double 'declares variable
Dim d As Double 'declares variable
Dim e As Double 'declares variable
Dim f As Double 'declares variable
Dim g As Double 'declares variable
Dim h As Double 'declares variable
Dim i As Double 'declares variable
Dim j As Double 'declares variable
Dim k As Double 'declares variable
Dim l As Double 'declares variable
Dim m As Double 'declares variable
Dim n As Double 'declares variable
Dim tpay As Double 'declares variable
Dim taxableincome As Double 'declares variable
Dim nicont As Double 'declares variable
Dim taxinc As Double 'declares variable
Dim tall As Double 'declares variable
Dim mainall As Double 'declares variable
Dim aaa As Double 'declares variable
Dim bb As Double 'declares variable
Dim cc As Double 'declares variable
Dim dd As Double 'declares variable
Dim ee As Double 'declares variable
Dim ff As Double 'declares variable
Dim npp As Double 'declares variable
dd = 0 'sets the value of dd to 0
bb = 0 'sets the value of bb to 0
ff = 0 'sets the value of ff to 0

If Not IsNumeric(Text5(0).Text) Then   'starts if statement
   MsgBox "Please enter proper value in age box", vbInformation 'shows message box
   Exit Sub 'exits sub
   
End If 'closes end if statement












mainall = 0 'sets the value of mainall to 0

Open App.Path + "\taxrates.dLL" For Input As #aa 'opens the file
Input #aa, a, b, c, d, e, f, g, h, i, j, k, l, m, n 'puts the values from file to declare variables
Close #aa 'closes file
Form4.Text1.Text = a 'set the value of the variable to particular text box of the form
Form4.Text2.Text = b 'set the value of the variable to particular text box of the form
Form4.Text3.Text = c 'set the value of the variable to particular text box of the form
Form4.Text4.Text = d 'set the value of the variable to particular text box of the form
Form4.Text5.Text = e 'set the value of the variable to particular text box of the form
Form4.Text6.Text = f 'set the value of the variable to particular text box of the form
Form4.Text7.Text = g 'set the value of the variable to particular text box of the form
Form4.Text8.Text = h 'set the value of the variable to particular text box of the form
Form4.Text9.Text = i 'set the value of the variable to particular text box of the form
Form4.Text10.Text = j 'set the value of the variable to particular text box of the form
Form4.Text11.Text = k 'set the value of the variable to particular text box of the form
Form4.Text12.Text = l 'set the value of the variable to particular text box of the form
Form4.Text13.Text = m 'set the value of the variable to particular text box of the form
Form4.Text14.Text = n 'set the value of the variable to particular text box of the form
Label3(6) = (Label3(10) * Form2.Text1.Text) + overtime
'counts total pay
If Combo5.Text = "Yes" Then 'counts allowances
   mainall = mainall + Form4.Text14.Text 'sets the value of the declared variable
End If ' end the endif function
    If Combo3.Text = "Yes" Then 'starts if statement
      mainall = mainall + Form4.Text13.Text
    End If
       If Combo4.Text = "Yes" Then 'starts if statement
        mainall = mainall + Form4.Text11.Text
        End If
                       If Text5(0).Text >= 40 Then 'starts if statement
                         mainall = mainall + Form4.Text12.Text
                       End If
                       Label3(5) = mainall
If Label3(5) > Label3(6) Then 'counts taxable income and starts if statement
 Label3(7) = 0 'sets the label value
 Else
   Label3(7) = Label3(6) - Label3(5) 'sets the label value
End If
If Label3(7) < 70 Then 'counts NI number and tax
    Label3(8) = Form4.Text5.Text * Label3(7) 'sets the label value
    Label3(0) = Form4.Text1.Text * Label3(7) 'sets the label value
    ElseIf Label3(7) < 114 Then
     Label3(8) = Form4.Text6.Text * Label3(7) 'sets the label value
       Label3(0) = (Form4.Text1.Text * Label3(7)) + (Form4.Text2.Text * Label3(7)) 'sets the label value
     ElseIf Label3(7) < 500 Then
     Label3(8) = Form4.Text7.Text * Label3(7) 'sets the label value
     Label3(0) = (Form4.Text1.Text * Label3(7)) + (Form4.Text2.Text * Label3(7)) + (Form4.Text3.Text * Label3(7)) 'sets the label value
     ElseIf Label3(7) > 500 Then
     Label3(8) = Form4.Text8.Text * Label3(7) 'sets the label value
      Label3(0) = (Form4.Text1.Text * Label3(7)) + (Form4.Text2.Text * Label3(7)) + (Form4.Text3.Text * Label3(7)) + (Form4.Text4.Text * Label3(7)) 'sets the label value
End If
nicont = Label3(8) 'counts net pay
taxinc = Label3(0) 'sets the value of the variable
DEDUCTIONS = nicont + taxin 'sets the value of the variable
If Check1 Then 'starts if statement
 DEDUCTIONS = DEDUCTIONS + (Label3(6) * Form4.Text9.Text) 'sets the value of the variable
End If
If Check2 Then 'starts if statement
 DEDUCTIONS = DEDUCTIONS + Form4.Text10.Text 'sets the value of the variable
End If
tpay = Label3(6) 'sets the value of the variable
taxableincome = DEDUCTIONS 'sets the value of the variable
npp = tpay - taxableincome 'sets the value of the variable
Label3(4) = npp 'sets the value of the label box
Label3(1) = Val(Label3(1)) + Val(Label3(8)) 'counts total ni contribution
Label3(2) = Val(Label3(2)) + Val(Label3(0)) 'counts total tax deducted so far
Label3(3) = Val(Label3(3)) + Val(Label3(4)) 'counts total earnings




Frame1.Visible = True 'make the frame visible
Command2.Visible = True 'makes the button visible
Form3.Width = 11835 'set the form width
End Sub

Private Sub Command2_Click()
Frame1.Visible = False 'hides the frame
Form3.Width = 8445 'set the form width

End Sub

Private Sub Form_Load()
Form3.Width = 8445 'set the form width
If Form2!Combo1.Text = "W" Then 'starts if statement
  Label3(10) = "£8.00" 'sets the value of the label
ElseIf Form2!Combo1.Text = "S" Then
Label3(10) = "£6.00" 'sets the value of the label
ElseIf Form2!Combo1.Text = "U" Then
Label3(10) = "£4" 'sets the value of the label
End If
Combo1.Text = Form2!Combo1.Text 'sets the value of the combo box
Label3(9) = Date 'sets the value of the label
Label3(6) = Format((Label3(10) * Form2!Text1.Text), "£###,###.00") 'sets the value of the label
Combo5.AddItem "Yes" 'sets the value of the combo box
Combo5.AddItem "No" 'sets the value of the combo box
Combo2.AddItem "M" 'sets the value of the combo box
Combo2.AddItem "F" 'sets the value of the combo box
Combo4.AddItem "Yes" 'sets the value of the combo box
Combo4.AddItem "No" 'sets the value of the combo box
Combo3.AddItem "Yes" 'sets the value of the combo box
Combo3.AddItem "No" 'sets the value of the combo box

End Sub

Private Sub Frame2_Click()
Frame1.Visible = False 'makes the frame invisible
Form3.Width = 8445 'sets the width of the form
End Sub

Private Sub save_Click()
Text4(1).Text = Date 'sets the value of the text box

aa = FreeFile 'sets the variable aa to freefile
Open App.Path + "\payroll.dLL" For Append As #aa 'opens the file
Print #aa, Text1(0), Chr(34); Text2(0).Text; Chr(34); , Chr(34); Text3(0).Text; Chr(34); , Chr(34); Combo1; Chr(34); , Chr(34); Combo5; Chr(34); , Chr(34); Text5(0); Chr(34); , Chr(34); Label3(10); Chr(34); , Chr(34); Combo2; Chr(34); , Chr(34); Combo3; Chr(34); , Chr(34); Combo4; Chr(34); , Chr(34); Text10(0); Chr(34); , Chr(34); Text1(1); Chr(34); , Chr(34); Text2(1); Chr(34); , Chr(34); Label3(9); Chr(34); , Chr(34); Text4(1); Chr(34); , Chr(34); Label3(6); Chr(34); , Chr(34); Label3(5); Chr(34); , Chr(34); Label3(7); Chr(34); , Chr(34); Label3(8); Chr(34); , Chr(34); Label3(0); Chr(34); , Chr(34); Label3(1); Chr(34); , Chr(34); Label3(2); Chr(34); , Chr(34); Label3(3); Chr(34); , Chr(34); Label3(4); Chr(34) 'writes to the file
Close #aa 'closes the file
End Sub

Private Sub Text1_Change(Index As Integer)
If Not IsNumeric(Text1(0).Text) Then  'starts if statement
   MsgBox "Please enter numeric value in employee Id", vbInformation 'shows message box
   Exit Sub 'exits sub
   
End If 'closes end if statement
End Sub

Private Sub Text2_Change(Index As Integer)
If IsNumeric(Text2(0).Text) Then   'starts if statement
   MsgBox "Please enter proper value in Name box", vbInformation 'shows message box
   Exit Sub 'exits sub
   
End If 'closes end if statement
End Sub

Private Sub Text3_Change(Index As Integer)
If IsNumeric(Text3(0).Text) Then   'starts if statement
   MsgBox "Please enter proper value in Surname box", vbInformation 'shows message box
   Exit Sub 'exits sub
   
End If 'closes end if statement
End Sub

Private Sub Text5_Change(Index As Integer)
If Not IsNumeric(Text5(0).Text) Then   'starts if statement
   MsgBox "Please enter proper value in age box", vbInformation 'shows message box
   Exit Sub 'exits sub
   
End If 'closes end if statement
End Sub

Private Sub ww_Click(Index As Integer)
Form1.Show 'shows the form
Form3.Hide 'hides the form
End Sub
