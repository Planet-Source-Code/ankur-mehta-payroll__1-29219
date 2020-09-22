VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Professional Payroll software "
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   960
   ClientWidth     =   11880
   Icon            =   "payslip4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7095
      Left            =   2160
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   12515
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   15360
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "payslip4.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   1
         Left            =   1440
         Picture         =   "payslip4.frx":1275
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton ee 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   2
         Left            =   2760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "payslip4.frx":1B1D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   3
         Left            =   4080
         Picture         =   "payslip4.frx":25DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   4
         Left            =   5400
         Picture         =   "payslip4.frx":30DD
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   5
         Left            =   6720
         Picture         =   "payslip4.frx":3AB2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   6
         Left            =   8040
         Picture         =   "payslip4.frx":44FE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton aazcxvccvf 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   7
         Left            =   9360
         Picture         =   "payslip4.frx":51E2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   615
         Index           =   8
         Left            =   10680
         Picture         =   "payslip4.frx":5AF2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1695
         Begin VB.CommandButton Command10 
            Height          =   615
            Left            =   360
            TabIndex        =   26
            Top             =   5040
            Width           =   1095
         End
         Begin VB.CommandButton Command9 
            Height          =   615
            Left            =   360
            TabIndex        =   25
            Top             =   4320
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Height          =   615
            Left            =   360
            TabIndex        =   24
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Height          =   735
            Left            =   360
            TabIndex        =   23
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Height          =   735
            Left            =   360
            TabIndex        =   22
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Height          =   735
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   735
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CommandButton Command14 
            Height          =   975
            Left            =   240
            TabIndex        =   18
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Height          =   975
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CommandButton Command12 
            Height          =   975
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command11 
            Height          =   975
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Modules"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Design"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid label1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   100
      Cols            =   25
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      TextStyle       =   4
      GridLines       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu q 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu w 
         Caption         =   "Employee Wizard"
         Index           =   2
      End
      Begin VB.Menu e 
         Caption         =   "Upload records"
         Index           =   3
      End
      Begin VB.Menu r 
         Caption         =   "Backup"
         Index           =   4
      End
      Begin VB.Menu a 
         Caption         =   "Exit"
         Index           =   6
      End
   End
   Begin VB.Menu v 
      Caption         =   "View "
      Index           =   9
      Begin VB.Menu h 
         Caption         =   "Toolbar"
         Index           =   12
      End
      Begin VB.Menu cv 
         Caption         =   "Backup"
         Index           =   56
      End
   End
   Begin VB.Menu modules 
      Caption         =   "Modules"
      Begin VB.Menu ddddd 
         Caption         =   "Employees"
      End
   End
   Begin VB.Menu hjkl 
      Caption         =   "Tasks"
      Index           =   122
      Begin VB.Menu ene 
         Caption         =   "Enter new employee"
         Index           =   1234
      End
      Begin VB.Menu ses 
         Caption         =   "Show employee statistics"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu tr 
         Caption         =   "Tax rates"
         Index           =   12345
      End
      Begin VB.Menu entr 
         Caption         =   "Enter new tax rates"
         Index           =   12345
      End
      Begin VB.Menu s 
         Caption         =   "Search"
         Index           =   1231
      End
   End
   Begin VB.Menu sdfghjkll 
      Caption         =   "Help"
      Begin VB.Menu hp 
         Caption         =   "Help topics"
         Index           =   181
      End
      Begin VB.Menu au 
         Caption         =   "About us"
         Index           =   161
      End
   End
End
Attribute VB_Name = "Form1"
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
Private Sub a_Click(Index As Integer)

End 'exits the program
End Sub

Private Sub aazcxvccvf_Click(Index As Integer)
End 'exits the program
End Sub

Private Sub Command2_Click()
If Frame3.Visible = False Then 'sets the boolean variable
  Frame3.Visible = True 'shows the frame
  Frame4.Visible = False 'hides the frame
  
  ElseIf Frame3.Visible = True Then 'starts the else if statement
   Frame3.Visible = False 'hides the form
   Frame4.Visible = True 'shows the form
   
   End If 'closes the if statement
End Sub

Private Sub Command3_Click()
If Frame3.Visible = False Then 'sets the boolean variable
  Frame3.Visible = True 'shows the frame
  Frame4.Visible = False 'hides the frame
  
  ElseIf Frame3.Visible = True Then 'sets the boolean variable
   Frame3.Visible = False 'hides the frame
   Frame4.Visible = True 'showes the frame
   
   End If 'ends the endif statement
End Sub

Private Sub cv_Click(Index As Integer)
Dim fys As Object
Set fsys = CreateObject("scripting.filesystemobject")
fsys.copyfile App.Path + "\PAYROLL.DLL", "c:\WINDOWS\"
fsys.copyfile App.Path + "\TAXRATES.DLL", "c:\WINDOWS\"
End Sub

Private Sub ddddd_Click()
Call main

End Sub

Private Sub e_Click(Index As Integer)
On Error GoTo errorhandler 'handles any sort of error
Dim X As Integer 'declares variable
Dim xx As Integer 'declares variable
Dim aa As String 'declares variable

 
 
aa = FreeFile 'set the variable aa to freefile


Open App.Path + "\payroll.dLL" For Input As aa 'opens the file
For X = 1 To 100 'starts the inrestic control
  Input #aa, id(X), fname(X), sname(X), jtitle(X), disable(X), age(X), payrate(X), gender(X), widow(X), married(X), ninumber(X), address(X), telephone(X), cpaydate(X), lpaydate(X), Totalpay(X), allowance(X), taxableincome(X), Nicontribution(X), tax(X), totalnicontribution(X), totaltax(X), totalearning(X), netpay(X) 'set the values from the file to specific variables
Label1.TextMatrix(X, 1) = id(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 2) = fname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 3) = sname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 4) = jtitle(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 5) = disable(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 6) = age(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 7) = payrate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 8) = gender(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 9) = widow(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 10) = married(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 11) = ninumber(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 12) = address(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 13) = telephone(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 14) = cpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 15) = lpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 16) = Totalpay(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 17) = allowance(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 18) = taxableincome(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 19) = Nicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 20) = tax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 21) = totalnicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 22) = totaltax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 23) = totalearning(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 24) = netpay(X) ' puts the variable values in flexgrid
Next X 'starts the loop again

Close #aa 'closes the file
If Frame1.Visible = True Then 'sets boolean variable


Label1.Top = 720 'sets the flexgrid top
Else

Label1.Top = 0 'sets the flexgrid top

End If 'ends the end if statement

ses(10).Enabled = True 'makes the button enable true

errorhandler:
        Debug.Print "error" 'prints the debug message
        Close #aa 'closes file
        
 
End Sub

Private Sub ee_Click(Index As Integer)

On Error GoTo errorhandler 'handles any sort of error
Dim X As Integer 'declares variable
Dim xx As Integer 'declares variable
Dim aa As String 'declares variable

 
 
aa = FreeFile 'set the variable aa to freefile


Open App.Path + "\payroll.dLL" For Input As aa 'opens the file
For X = 1 To 100 'starts the inrestic control
  Input #aa, id(X), fname(X), sname(X), jtitle(X), disable(X), age(X), payrate(X), gender(X), widow(X), married(X), ninumber(X), address(X), telephone(X), cpaydate(X), lpaydate(X), Totalpay(X), allowance(X), taxableincome(X), Nicontribution(X), tax(X), totalnicontribution(X), totaltax(X), totalearning(X), netpay(X) 'set the values from the file to specific variables
Label1.TextMatrix(X, 1) = id(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 2) = fname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 3) = sname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 4) = jtitle(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 5) = disable(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 6) = age(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 7) = payrate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 8) = gender(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 9) = widow(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 10) = married(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 11) = ninumber(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 12) = address(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 13) = telephone(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 14) = cpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 15) = lpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 16) = Totalpay(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 17) = allowance(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 18) = taxableincome(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 19) = Nicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 20) = tax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 21) = totalnicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 22) = totaltax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 23) = totalearning(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 24) = netpay(X) ' puts the variable values in flexgrid
Next X 'starts the loop again

Close #aa 'closes the file
If Frame1.Visible = True Then 'sets boolean variable


Label1.Top = 720 'sets the flexgrid top
Else

Label1.Top = 0 'sets the flexgrid top

End If 'ends the end if statement

ses(10).Enabled = True 'makes the button enable true

errorhandler:
        Debug.Print "error" 'prints the debug message
        Close #aa 'closes file
        
 
End Sub

Private Sub ene_Click(Index As Integer)
Form2.Show 'shows the form


End Sub

Private Sub entr_Click(Index As Integer)

Dim aa As Integer 'declares the variable
aa = FreeFile 'sets the variable to aa
Dim a As Double 'declares the variable
Dim b As Double 'declares the variable
Dim c As Double 'declares the variable
Dim d As Double 'declares the variable
Dim e As Double 'declares the variable
Dim f As Double 'declares the variable
Dim g As Double 'declares the variable
Dim h As Double 'declares the variable
Dim i As Double 'declares the variable
Dim j As Double 'declares the variable
Dim k As Double 'declares the variable
Dim l As Double 'declares the variable
Dim m As Double 'declares the variable
Dim n As Double 'declares the variable

Open App.Path + "\taxrates.dLL" For Input As #aa 'opens the file
Input #aa, a, b, c, d, e, f, g, h, i, j, k, l, m, n 'input the values from the file to these declares variables

Close #aa 'close the file
Form4.Text1.Text = a 'sets the value of the text box to particular form
Form4.Text2.Text = b 'sets the value of the text box to particular form
Form4.Text3.Text = c 'sets the value of the text box to particular form
Form4.Text4.Text = d 'sets the value of the text box to particular form
Form4.Text5.Text = e 'sets the value of the text box to particular form
Form4.Text6.Text = f 'sets the value of the text box to particular form
Form4.Text7.Text = g 'sets the value of the text box to particular form
Form4.Text8.Text = h 'sets the value of the text box to particular form
Form4.Text9.Text = i 'sets the value of the text box to particular form
Form4.Text10.Text = j 'sets the value of the text box to particular form
Form4.Text11.Text = k 'sets the value of the text box to particular form
Form4.Text12.Text = l 'sets the value of the text box to particular form
Form4.Text13.Text = m 'sets the value of the text box to particular form
Form4.Text14.Text = n 'sets the value of the text box to particular form
Form4.Show 'shows the form
End Sub

Private Sub Form_Activate()
Call main 'calls the main procedure
End Sub

Private Sub Form_GotFocus()
Call main 'calls the main procedure
End Sub

Private Sub Form_Load()
Dim X As Integer 'declares the variable
Dim a As Node
Set a = TreeView1.Nodes.Add(, , "7 node", "Payroll Software")

Set a = TreeView1.Nodes.Add("7 node", tvwChild, "c", "Customers")
Set a = TreeView1.Nodes.Add("7 node", tvwChild, "e", "Employees")
Set a = TreeView1.Nodes.Add("7 node", tvwChild, "j", "Job Available")
Set a = TreeView1.Nodes.Add("7 node", tvwChild, "s", "Suppliers")
Set a = TreeView1.Nodes.Add("7 node", tvwChild, "st", "Stock management")

Set a = TreeView1.Nodes.Add("c", tvwChild, , "Permanent Customers")
Set a = TreeView1.Nodes.Add("c", tvwChild, , "Part Time Customers")



Set a = TreeView1.Nodes.Add("e", tvwChild, , "Full Time employees")
Set a = TreeView1.Nodes.Add("e", tvwChild, , "Part Time employees")

Set a = TreeView1.Nodes.Add("j", tvwChild, , "Permanent Jobs")
Set a = TreeView1.Nodes.Add("j", tvwChild, , "Part Time Jobs")

Set a = TreeView1.Nodes.Add("s", tvwChild, , "Permanent Suppliers")
Set a = TreeView1.Nodes.Add("s", tvwChild, , "Part Time Suppliers")


Set a = TreeView1.Nodes.Add("st", tvwChild, , "Stock on selfs")
Set a = TreeView1.Nodes.Add("st", tvwChild, , "Stock on order")
Set a = TreeView1.Nodes.Add("st", tvwChild, , "Stock received")





Frame1.Width = Me.Width + 222 'sets the frame width
If Frame1.Visible = True And Label1.Visible = True Then 'starts the if statement
Label1.Top = 0 'sets the top of the flex grid
Frame1.Top = 0 'sets the top of the frame
Else
 Frame1.Top = 0 'sets the top of the frame
End If
Label1.TextMatrix(0, 1) = "Employee ID" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 2) = "Name" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 3) = "Surname" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 4) = "Job Title" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 5) = "Disable" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 6) = "Age" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 7) = "Payrate" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 8) = "Gender" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 9) = "Widow" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 10) = "Married" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 11) = "NI number" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 12) = "Address" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 13) = "Telephone" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 14) = "Current paydate" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 15) = "Last paydate" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 16) = "Total pay" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 17) = "Total allowance" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 18) = "Taxable income" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 19) = "NI contribution" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 20) = "Tax" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 21) = "Total NI contribution" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 22) = "Total Tax" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 23) = "Total aearning" 'sets the value of the flex grid cell to particular variable
Label1.TextMatrix(0, 24) = "Net pay" 'sets the value of the flex grid cell to particular variable


For X = 1 To 99 'starts the inrestic control
  Label1.TextMatrix(X, 0) = X 'set the value of flexgrid cell to x
Next 'loops








 


End Sub

Private Sub Form_Resize()
Frame2.Height = Me.Height
Frame3.Height = Me.Height

TreeView1.Height = Me.Height - 1500
Label1.Width = Form1.Width - (100)
Label1.Height = Form1.Height - 1500
End Sub

Private Sub h_Click(Index As Integer)
If Frame1.Visible = True And Label1.Visible = True Then 'starts the boolean variable

Frame1.Visible = False 'hides the frame
Label1.Top = 0 'sets the flexgid top
Frame2.Top = 0 'sets the top of the frame
Else
Frame1.Visible = True 'shows the frame
Label1.Top = 720 'sets the flexgid top
Frame2.Top = 720 'sets the top of the frame
End If 'end the endif statement
End Sub



Private Sub label1_DblClick()
Dim a As Integer 'declares the variable
a = Label1.Row 'sets the value of the declared variable
Form5.Show 'shows the form
    Form5.Text1(0) = id(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text2(0) = fname(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text3(0) = sname(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Combo1 = jtitle(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Combo5 = disable(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text5(0) = age(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(10) = payrate(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Combo2 = gender(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Combo3 = widow(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Combo4 = married(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text10(0) = ninumber(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text1(1) = address(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Text2(1) = telephone(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(9) = cpaydate(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.text48 = lpaydate(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(6) = Totalpay(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(5) = allowance(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(7) = taxableincome(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(8) = Nicontribution(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(0) = tax(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(1) = totalnicontribution(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(2) = totaltax(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(3) = totalearning(Label1.Row) 'sets the text box value to the particular flex grid cell text
    Form5.Label3(4) = netpay(Label1.Row) 'sets the text box value to the particular flex grid cell text
    save = (Label1.Row) 'declares the variable value
End Sub

Private Sub r_Click(Index As Integer)
Set fsys = CreateObject("scripting.filesystemobject")
fsys.copyfile App.Path + "\PAYROLL.DLL", "c:\WINDOWS\"
fsys.copyfile App.Path + "\TAXRATES.DLL", "c:\WINDOWS\"
End Sub

Private Sub s_Click(Index As Integer)

Form6.Show 'shows the form
End Sub



Private Sub tr_Click(Index As Integer)
Form4.Show 'shows the form
Form1.Hide 'hides the form
Dim aa As Integer 'declares the variable
aa = FreeFile 'sets the aa variable to freefile
Dim a As Double 'declares the variable
Dim b As Double 'declares the variable
Dim c As Double 'declares the variable
Dim d As Double 'declares the variable
Dim e As Double 'declares the variable
Dim f As Double 'declares the variable
Dim g As Double 'declares the variable
Dim h As Double 'declares the variable
Dim i As Double 'declares the variable
Dim j As Double 'declares the variable
Dim k As Double 'declares the variable
Dim l As Double 'declares the variable
Dim m As Double 'declares the variable
Dim n As Double 'declares the variable

Open App.Path + "\taxrates.dLL" For Input As #aa 'opens the file
Input #aa, a, b, c, d, e, f, g, h, i, j, k, l, m, n 'set the variable to the data reterive from the file

Close #aa 'closes the file
Form4.Text1.Text = a 'sets the value of the textbox of the particular form to declared variable
Form4.Text2.Text = b 'sets the value of the textbox of the particular form to declared variable
Form4.Text3.Text = c 'sets the value of the textbox of the particular form to declared variable
Form4.Text4.Text = d 'sets the value of the textbox of the particular form to declared variable
Form4.Text5.Text = e 'sets the value of the textbox of the particular form to declared variable
Form4.Text6.Text = f 'sets the value of the textbox of the particular form to declared variable
Form4.Text7.Text = g 'sets the value of the textbox of the particular form to declared variable
Form4.Text8.Text = h 'sets the value of the textbox of the particular form to declared variable
Form4.Text9.Text = i 'sets the value of the textbox of the particular form to declared variable
Form4.Text10.Text = j 'sets the value of the textbox of the particular form to declared variable
Form4.Text11.Text = k 'sets the value of the textbox of the particular form to declared variable
Form4.Text12.Text = l 'sets the value of the textbox of the particular form to declared variable
Form4.Text13.Text = m 'sets the value of the textbox of the particular form to declared variable
Form4.Text14.Text = n 'sets the value of the textbox of the particular form to declared variable
Form4.Show 'shows the form
Form4.Text1.Locked = True 'locks the particular text box of the particular form
Form4.Text2.Locked = True 'locks the particular text box of the particular form
Form4.Text3.Locked = True 'locks the particular text box of the particular form
Form4.Text4.Locked = True 'locks the particular text box of the particular form
Form4.Text5.Locked = True 'locks the particular text box of the particular form
Form4.Text6.Locked = True 'locks the particular text box of the particular form
Form4.Text7.Locked = True 'locks the particular text box of the particular form
Form4.Text8.Locked = True 'locks the particular text box of the particular form
Form4.Text9.Locked = True 'locks the particular text box of the particular form
Form4.Text10.Locked = True 'locks the particular text box of the particular form
Form4.w.Enabled = False 'sets the enable of the w object of the particular form to false

End Sub

Private Sub w_Click(Index As Integer)
Form2.Show 'shows the form

End Sub
Sub main()

On Error GoTo errorhandler 'handles any sort of error
Dim X As Integer 'declares variable
Dim xx As Integer 'declares variable
Dim aa As String 'declares variable

 
 
aa = FreeFile 'set the variable aa to freefile


Open App.Path + "\payroll.dLL" For Input As aa 'opens the file
For X = 1 To 100 'starts the inrestic control
  Input #aa, id(X), fname(X), sname(X), jtitle(X), disable(X), age(X), payrate(X), gender(X), widow(X), married(X), ninumber(X), address(X), telephone(X), cpaydate(X), lpaydate(X), Totalpay(X), allowance(X), taxableincome(X), Nicontribution(X), tax(X), totalnicontribution(X), totaltax(X), totalearning(X), netpay(X) 'set the values from the file to specific variables
Label1.TextMatrix(X, 1) = id(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 2) = fname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 3) = sname(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 4) = jtitle(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 5) = disable(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 6) = age(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 7) = payrate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 8) = gender(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 9) = widow(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 10) = married(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 11) = ninumber(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 12) = address(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 13) = telephone(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 14) = cpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 15) = lpaydate(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 16) = Totalpay(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 17) = allowance(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 18) = taxableincome(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 19) = Nicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 20) = tax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 21) = totalnicontribution(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 22) = totaltax(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 23) = totalearning(X) ' puts the variable values in flexgrid
Label1.TextMatrix(X, 24) = netpay(X) ' puts the variable values in flexgrid
Next X 'starts the loop again

Close #aa 'closes the file
If Frame1.Visible = True Then 'sets boolean variable


Label1.Top = 720 'sets the flexgrid top
Else

Label1.Top = 0 'sets the flexgrid top

End If 'ends the end if statement

ses(10).Enabled = True 'makes the button enable true

errorhandler:
        Debug.Print "error" 'prints the debug message
        Close #aa 'closes file
        
 
End Sub
