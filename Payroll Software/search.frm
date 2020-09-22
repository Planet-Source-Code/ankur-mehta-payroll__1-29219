VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   795
   ClientLeft      =   3060
   ClientTop       =   3255
   ClientWidth     =   9480
   ForeColor       =   &H8000000D&
   Icon            =   "search.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   795
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search For customer"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter employee id number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter employee NI number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please select the toplic of the search"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id(1 To 10) As Integer 'declares the variable with arrays


Dim fname(1 To 10) As String 'declares the variable with arrays



Dim sname(1 To 10) As String 'declares the variable with arrays
Dim jtitle(1 To 10) As String 'declares the variable with arrays


Dim disable(1 To 10) 'declares the variable with arrays
Dim age(1 To 10) 'declares the variable with arrays

Dim payrate(10) 'declares the variable with arrays


Dim gender(1 To 10) 'declares the variable with arrays

Dim widow(1 To 10) 'declares the variable with arrays
Dim married(1 To 10) 'declares the variable with arrays
Dim ninumber(1 To 10) 'declares the variable with arrays

Dim address(1 To 10) 'declares the variable with arrays

Dim telephone(10) 'declares the variable with arrays

Dim cpaydate(1 To 10) 'declares the variable with arrays
Dim lpaydate(10) 'declares the variable with arrays
Dim Totalpay(1 To 10) 'declares the variable with arrays
Dim allowance(1 To 10) 'declares the variable with arrays
Dim taxableincome(1 To 10) 'declares the variable with arrays
Dim Nicontribution(1 To 10) 'declares the variable with arrays
Dim tax(1 To 10) 'declares the variable with arrays
Dim totalnicontribution(1 To 10) 'declares the variable with arrays
Dim totaltax(1 To 10) 'declares the variable with arrays
Dim totalearning(1 To 10) 'declares the variable with arrays
Dim netpay(1 To 10) 'declares the variable with arrays
Dim result As Boolean 'declares boolean variable
Private Sub Command1_Click()
Dim result As Boolean 'declares boolean variable
Dim sid As Integer 'declares variable
Dim X As Integer 'declares variable
Dim Y As Integer 'declares variable
Dim aa As String 'declares variable
aa = FreeFile 'sets aa variable as freefile
Open App.Path + "\payroll.dLL" For Input As aa 'opens the file
For X = 1 To 1 'starts the inrestic control
  Input #aa, id(X), fname(X), sname(X), jtitle(X), disable(X), age(X), payrate(X), gender(X), widow(X), married(X), ninumber(X), address(X), telephone(X), cpaydate(X), lpaydate(X), Totalpay(X), allowance(X), taxableincome(X), Nicontribution(X), tax(X), totalnicontribution(X), totaltax(X), totalearning(X), netpay(X) 'set the values from the file to specific variables
Next X 'starts the loop again

Close #aa 'closes the file
If Text1.Text <> "" Then 'sets the boolean operator
  Call one 'calls procedure
    ElseIf Text2.Text <> "" Then 'sets the boolean operator
      Call Secondone 'calls procedure
        ElseIf Text1.Text <> "" And Text2.Text <> "" Then 'sets the boolean operator
        Call lastone 'calls procedure

            ElseIf result = True Then 'sets the boolean operator
             Exit Sub ' exits the sub
                ElseIf result = False Then 'sets the result to false
                MsgBox "Sorry no record find according to your information", vbCritical, "Search Results" 'displays message box
End If 'ends the end if
End Sub

Sub one()
Dim X As Integer 'declares variable
For X = 1 To 1 ' starts for loop
  If id(X) = Text1.Text Then
                Form5.Text1(0) = id(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(0) = fname(1) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text3(0) = sname(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo1 = jtitle(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo5 = disable(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text5(0) = age(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(10) = payrate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo2 = gender(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo3 = widow(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo4 = married(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text10(0) = ninumber(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text1(1) = address(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(1) = telephone(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(9) = cpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.text48 = lpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(6) = Totalpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(5) = allowance(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(7) = taxableincome(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(8) = Nicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(0) = tax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(1) = totalnicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(2) = totaltax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(3) = totalearning(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(4) = netpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form6.Hide 'hides the form
                Form5.Show 'shows the form
                 result = True 'sets result variable to true
                Exit Sub
  End If
  Next X
End Sub
Sub Secondone()
Dim X As Integer 'declares variable
For X = 1 To 1 ' starts for loop
  If ninumber(X) = Text2.Text Then
                Form5.Text1(0) = id(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(0) = fname(1) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text3(0) = sname(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo1 = jtitle(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo5 = disable(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text5(0) = age(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(10) = payrate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo2 = gender(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo3 = widow(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo4 = married(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text10(0) = ninumber(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text1(1) = address(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(1) = telephone(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(9) = cpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.text48 = lpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(6) = Totalpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(5) = allowance(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(7) = taxableincome(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(8) = Nicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(0) = tax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(1) = totalnicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(2) = totaltax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(3) = totalearning(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(4) = netpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form6.Hide 'hides the form
                Form5.Show 'shows the form
                 result = True 'sets result variable to true
                Exit Sub
  End If
  Next X
End Sub
Sub lastone()
Dim X As Integer 'declares variable
For X = 1 To 1 ' starts for loop
  If ninumber(X) = Text2.Text And id(X) = Text1.Text Then
                         Form5.Text1(0) = id(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(0) = fname(1) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text3(0) = sname(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo1 = jtitle(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo5 = disable(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text5(0) = age(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(10) = payrate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo2 = gender(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo3 = widow(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Combo4 = married(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text10(0) = ninumber(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text1(1) = address(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Text2(1) = telephone(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(9) = cpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.text48 = lpaydate(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(6) = Totalpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(5) = allowance(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(7) = taxableincome(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(8) = Nicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(0) = tax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(1) = totalnicontribution(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(2) = totaltax(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(3) = totalearning(X) 'sets the value of the particular text box of the particular form to declared variable
                Form5.Label3(4) = netpay(X) 'sets the value of the particular text box of the particular form to declared variable
                Form6.Hide 'hides the form
                Form5.Show 'shows the form
                 result = True 'sets result variable to true
                Exit Sub
  End If
  Next X
End Sub

Private Sub Command2_Click()
Form6.Hide 'hides the form


End Sub

