VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Attendence 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11400
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   22125
   LinkTopic       =   "Form1"
   Picture         =   "Attendence.frx":0000
   ScaleHeight     =   11400
   ScaleWidth      =   22125
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Attendance"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16560
      TabIndex        =   11
      Top             =   6840
      Width           =   4935
      Begin VB.OptionButton Absent 
         BackColor       =   &H00FFFF00&
         Caption         =   "Absent"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Present 
         BackColor       =   &H00FFFF00&
         Caption         =   "Present"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComCtl2.DTPicker ADATE 
      Height          =   615
      Left            =   16440
      TabIndex        =   9
      Top             =   3120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109248512
      CurrentDate     =   43246
   End
   Begin VB.TextBox ENAME 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   18480
      TabIndex        =   6
      Top             =   6000
      Width           =   3015
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
      Height          =   612
      Left            =   18480
      TabIndex        =   5
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton ASTART 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start Attendence"
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   5055
   End
   Begin VB.CommandButton NEMP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Next"
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
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox SEMP 
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   6360
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   11218
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   6375
      Left            =   9360
      TabIndex        =   14
      Top             =   2640
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11245
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
      Left            =   16440
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "search :"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
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
      Height          =   615
      Left            =   16560
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
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
      Height          =   615
      Left            =   16560
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "Attendence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim saved As Boolean
Dim inc As Integer
Dim inc2 As Integer
Dim ecount As Boolean

Private Sub EUPDATE_Click()
If ADD.Caption = "Add" Then
End If
End Sub

Private Sub ASTART_Click()
inc2 = 1
Do While (inc2 <= ListView1.ListItems.Count)
ListView1.ListItems(inc2).selected = True
If ListView1.SelectedItem = ADATE.Value Then
ecount = True
End If
inc2 = inc2 + 1
Loop

If ecount = True Then
MsgBox "Attendence is already done for today"
ecount = False
Else
If Not ListView2.ListItems.Count = 0 Then
ecount = False
ADATE.Enabled = False
Present.Enabled = True
Absent.Enabled = True
NEMP.Enabled = True
inc = 1
ListView2.ListItems(inc).selected = True
ListView2.SetFocus
Else
MsgBox "No employees to add records"
End If
End If

End Sub


Private Sub Form_Load()
With ListView1.ColumnHeaders
.ADD , , "Date", 2000
.ADD , , "EmpID", 1500
.ADD , , "Name", 2000
.ADD , , "Status", 1500
.ADD , , "Month", 0
End With
With ListView2.ColumnHeaders
.ADD , , "EmpID", 1500
.ADD , , "Name", 1500
End With
loaddata
EID.Enabled = False
Ename.Enabled = False
Present.Enabled = False
Absent.Enabled = False
NEMP.Enabled = False
ADATE.MaxDate = Now
ADATE.Value = Now

End Sub
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Sub loaddata()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from Attendence", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!date)
list.SubItems(1) = rs!EmpID
list.SubItems(2) = rs!Name
list.SubItems(3) = rs!Status
list.SubItems(4) = rs!Month
rs.MoveNext
Loop
rs.Close
ListView2.ListItems.Clear
rs.Open "select * from Employee", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView2.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
rs.MoveNext
Loop
rs.Close
End Sub
Private Sub scrBlue_change()
Picture1.BackColor = RGB(scrred.Value, scrgreen.Value, scrblue.Value)
End Sub





Private Sub Form_Unload(Cancel As Integer)
Menu.Show
connect.Close
End Sub

Private Sub ListView2_GotFocus()
EID.Text = ListView2.SelectedItem
Ename.Text = ListView2.SelectedItem.SubItems(1)
End Sub

Private Sub ListView2_Click()
ListView2.Enabled = False
End Sub

Private Sub NEMP_Click()

If Present.Value = False And Absent.Value = False Then
MsgBox "Please select Present or Absent!", , "Message"
Else

rs.Open "select * from Attendence", connect, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("AID") = EID + CStr(ADATE.Month) + CStr(ADATE.Day) + CStr(ADATE.Year)
rs.Fields("Date") = ADATE.Value
rs.Fields("EmpID") = EID.Text
rs.Fields("Name") = Ename.Text
If Present.Value = True Then
rs.Fields("Status") = "Present"
Else
rs.Fields("Status") = "Absent"
End If
rs.Fields("Month") = ADATE.Month
rs.Update
ListView1.ListItems.Clear
connect.Close
loaddata
Present.Value = False
Absent.Value = False
If inc < ListView2.ListItems.Count Then
inc = inc + 1
ListView2.ListItems(inc).selected = True
ListView2.Enabled = True
ListView2.SetFocus
Else
NEMP.Enabled = False
ADATE.Enabled = True
Present.Enabled = False
Absent.Enabled = False
NEMP.Enabled = False
MsgBox "attendence is over for today"
End If

End If
End Sub


Private Sub SEMP_Change()
ListView1.ListItems.Clear
rs.Open "select * from Attendence where EmpID like '" & SEMP & "%'", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!date)
list.SubItems(1) = rs!EmpID
list.SubItems(2) = rs!Name
list.SubItems(3) = rs!Status
list.SubItems(4) = rs!Month
rs.MoveNext
Loop
rs.Close
End Sub

