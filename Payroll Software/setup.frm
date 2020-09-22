VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim aa As Integer
Dim cc As String
Dim fsys As Object
Dim qq As String
Dim s As String
Dim w As String
Set fsys = CreateObject("scripting.filesystemobject")
qq = Text1.Text

If Text1.Text = "" Then
 s = MsgBox("Do you want to use default path settings", vbYesNo, "Path")
 If s = vbYes Then
    
    cc = "c:\payroll software\"
 ElseIf s = vbNo Then
   Exit Sub
   End If
Else
   cc = Text1.Text
End If

 
Form2.Visible = fasle
Form1.Command4.Enabled = True
path = cc

errorhandler:
  Debug.Print "ww"
  
  

End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.path
End Sub

Private Sub Drive1_Change()
Text1.Text = ""
Dir1.path = Drive1
Text1.Text = Dir1.path
End Sub

Private Sub Form_Load()
Me.Top = Form1.Height / 2 - 1500
Me.Left = Form1.Width / 2 - 3600
End Sub
