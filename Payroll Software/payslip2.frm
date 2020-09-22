VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2490
   ClientLeft      =   2055
   ClientTop       =   2010
   ClientWidth     =   7320
   ForeColor       =   &H00FFFFFF&
   Icon            =   "payslip2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7320
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hours worked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Job title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text > 40 Then
   overtimerate = InputBox("Please enter the overtime rate")
   overtime = (Text1.Text - 40) * overtimerate
End If

   

If Combo1.Text = "" Or Text1.Text = "" Or Text1.Text < 0 Then 'starts the endif statement
MsgBox "Syntex error try again", vbCritical, "Error" 'shows message box
Exit Sub 'exits the procedure
ElseIf Not IsNumeric(Text1.Text) Then
MsgBox "Syntex error try again", vbCritical, "Error" 'shows message box
Exit Sub 'exits the sub
Else
Form2.Hide 'hides form
Form3.Show 'shows form
End If 'closes endif statement

End Sub

Private Sub Command2_Click()
Form2.Hide 'hides form
Form1.Show 'shows form

End Sub

Private Sub Form_Load()
Combo1.AddItem "S" 'sets the value of the combo box
Combo1.AddItem "U" 'sets the value of the combo box
Combo1.AddItem "W" 'sets the value of the combo box
End Sub
