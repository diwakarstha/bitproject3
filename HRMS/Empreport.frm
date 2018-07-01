VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Empreport 
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21600
   BeginProperty Font 
      Name            =   "Copperplate Gothic Bold"
      Size            =   26.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Empreport.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   21600
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Ephone 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Ename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox EID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7815
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13785
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
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE REPORT"
      Height          =   1095
      Left            =   7680
      TabIndex        =   5
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15600
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "Empreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pc As Integer
Dim ac As Integer
Dim sm As Integer
Dim smc As Integer
Dim sp As Long
Dim empc As Integer

Private Sub Form_Load()
Me.Picture = LoadPicture("Pictures\entry.jpg")
With ListView1.ColumnHeaders
.ADD , , "EmpID", 1500
.ADD , , "Name", 2000
.ADD , , "Phone", 2000
.ADD , , "Present days", 2000
.ADD , , "Absent days", 2000
.ADD , , "Total Month paid", 2600
.ADD , , "Total salary paid", 2700
End With
EID.Visible = False
Ename.Visible = False
Ephone.Visible = False
ListView1.Enabled = False
date = DateTime.Now

empc = 1
empat.loaddata
If Not empat.ListView1.ListItems.Count = 0 Then
Do While empc <= empat.ListView1.ListItems.Count
empat.ListView1.ListItems(empc).selected = True
EID.Text = empat.ListView1.SelectedItem
Ename.Text = empat.ListView1.SelectedItem.SubItems(1)
Ephone.Text = empat.ListView1.SelectedItem.SubItems(2)
FAbsent.EID.Text = EID.Text

FAbsent.Estatus.Text = "Absent"
FAbsent.loaddata1
ac = FAbsent.ListView1.ListItems.Count

FAbsent.Estatus.Text = "Present"
FAbsent.loaddata1
pc = FAbsent.ListView1.ListItems.Count

salaryst.EID.Text = EID.Text
salaryst.Estatus.Text = "Paid"
salaryst.loaddata1
sm = salaryst.ListView1.ListItems.Count
sp = 0
smc = 1
Do While smc <= sm
salaryst.ListView1.ListItems(smc).selected = True
sp = sp + salaryst.ListView1.SelectedItem.SubItems(3)
smc = smc + 1
Loop
empc = empc + 1
Set list = ListView1.ListItems.ADD(, , EID.Text)
list.SubItems(1) = Ename.Text
list.SubItems(2) = Ephone.Text
list.SubItems(3) = pc
list.SubItems(4) = ac
list.SubItems(5) = sm
list.SubItems(6) = sp
Loop
Else
MsgBox "no report can be generated as there is no data"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Show
Empreport.Hide
End Sub
