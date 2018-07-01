VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Salary 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   FillColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   Picture         =   "Salary.frx":0000
   ScaleHeight     =   12075
   ScaleWidth      =   22800
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Pay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8160
      Width           =   5295
   End
   Begin VB.ComboBox Smonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   18720
      TabIndex        =   11
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Adays 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   512
      Left            =   18720
      TabIndex        =   10
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox Salary 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   512
      Left            =   18720
      TabIndex        =   9
      Top             =   6720
      Width           =   2895
   End
   Begin VB.TextBox Ename 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   512
      Left            =   18720
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox EID 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   512
      Left            =   18720
      TabIndex        =   3
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Ssalary 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6255
      Left            =   10680
      TabIndex        =   13
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6240
      Left            =   480
      TabIndex        =   14
      Top             =   2640
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11007
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker TDATE 
      Height          =   615
      Left            =   18720
      TabIndex        =   15
      Top             =   2640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109707265
      CurrentDate     =   43246
   End
   Begin VB.Label salarytxt 
      Alignment       =   2  'Center
      Caption         =   "NULL"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18720
      TabIndex        =   18
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Today's Date :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16320
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Absent days :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16320
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16320
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary of month :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16320
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   16320
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16320
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Emp ID :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16320
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "Salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim passmonth As Integer
Dim pc As Integer
Dim ac As Integer

Private Sub Form_Load()
Me.Picture = LoadPicture("Pictures\entry.jpg")
With ListView1.ColumnHeaders
.ADD , , "EmpID", 1000
.ADD , , "Name", 2000
.ADD , , "Salary Month", 2000
.ADD , , "Absent Days", 1500
.ADD , , "Salary", 1500
.ADD , , "Status", 2000
End With
With ListView2.ColumnHeaders
.ADD , , "EmpID", 1500
.ADD , , "Name", 2000
.ADD , , "Salary", 1500
End With
loaddata
TDATE.MaxDate = Now
TDATE.Value = Now

Smonth.AddItem 1
Smonth.AddItem 2
Smonth.AddItem 3
Smonth.AddItem 4
Smonth.AddItem 5
Smonth.AddItem 6
Smonth.AddItem 7
Smonth.AddItem 8
Smonth.AddItem 9
Smonth.AddItem 10
Smonth.AddItem 11
Smonth.AddItem 12

EID.Enabled = False
Ename.Enabled = False
Adays.Enabled = False
Salary.Enabled = False
Pay.Enabled = False
End Sub

Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Sub loaddata()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from Salary", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!Smonth
list.SubItems(3) = rs!Adays
list.SubItems(4) = rs!Salary
list.SubItems(5) = rs!Status
rs.MoveNext
Loop
rs.Close
ListView2.ListItems.Clear
rs.Open "select * from Employee", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView2.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!Salary
rs.MoveNext
Loop
rs.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
Menu.Show
connect.Close
End Sub

Private Sub ListView2_Click()
If Not ListView2.ListItems.Count = 0 Then
EID.Text = ListView2.SelectedItem
Ename.Text = ListView2.SelectedItem.SubItems(1)
Smonth.Text = Empty
Adays.Text = Empty
Salary.Text = Empty
salarytxt.BackColor = vbWhite
salarytxt.Caption = "NULL"
Else
MsgBox "No employees to pay salary"
End If
End Sub

Private Sub Pay_Click()
If salarytxt.Caption = "Unpaid" Then
rs.Open "select * from Salary", connect, adOpenDynamic, adLockOptimistic
rs.AddNew
'rs.Fields("SID") = EID.Text + Smonth.Text + CStr(TDATE.Year)
rs.Fields("EmpID") = EID.Text
rs.Fields("Name") = Ename.Text
rs.Fields("Smonth") = Smonth.Text
rs.Fields("Adays") = Adays.Text
rs.Fields("Salary") = Salary.Text
rs.Fields("Status") = "Paid"
rs.Update
ListView1.ListItems.Clear
connect.Close
loaddata
Pay.Enabled = False
End If

End Sub


Private Sub Smonth_Change()
If Smonth.Text = Empty Then
Smonth.Text = 1
Else
Dim mon As String
mon = Smonth.Text
If mon < 13 Then
Else
Smonth.Text = 1
End If
End If
End Sub

Private Sub Smonth_Click()
If Not EID.Text = Empty Then

FAbsent.EID.Text = EID.Text
FAbsent.Month.Text = CInt(Smonth.Text)
FAbsent.Estatus.Text = "Absent"
FAbsent.loaddata
ac = FAbsent.ListView1.ListItems.Count
Adays.Text = CStr(FAbsent.ListView1.ListItems.Count)
FAbsent.Estatus.Text = "Present"
FAbsent.loaddata
pc = FAbsent.ListView1.ListItems.Count
Salary.Text = CStr(FAbsent.ListView1.ListItems.Count) * (ListView2.SelectedItem.SubItems(2) / 30)

If (pc + ac > 29) Then
salaryst.EID.Text = EID.Text
salaryst.Month.Text = CInt(Smonth.Text)
salaryst.Estatus.Text = "Paid"
salaryst.loaddata
'salaryst.Show
'EID.Text = salaryst.ListView1.ListItems.Count
If salaryst.ListView1.ListItems.Count > 0 Then
salarytxt.BackColor = vbGreen
salarytxt.Caption = "Paid"
Pay.Enabled = False
Else
salarytxt.BackColor = vbRed
salarytxt.Caption = "Unpaid"
Pay.Enabled = True
End If
Else
salarytxt.BackColor = vbWhite
salarytxt.Caption = "NULL"
End If

Else
MsgBox "please select employee first"
End If
End Sub

Private Sub Smonth_KeyPress(KeyAscii As Integer)



If mon < 13 Then
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
Else
Smonth.Text = Empty
End If
End Sub

Private Sub Ssalary_Change()
ListView1.ListItems.Clear
rs.Open "select * from Salary where EmpID like '" & SEMP & "%'", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!Smonth
list.SubItems(3) = rs!Adays
list.SubItems(4) = rs!Salary
list.SubItems(5) = rs!Status
rs.MoveNext
Loop
rs.Close
End Sub

