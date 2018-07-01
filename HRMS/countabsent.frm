VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FAbsent 
   Caption         =   "Count Absent"
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   13020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Estatus 
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Month 
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox EID 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10821
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
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Month:"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "empid: "
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "FAbsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
With ListView1.ColumnHeaders
.ADD , , "Date", 2000
.ADD , , "EmpID", 1500
.ADD , , "Name", 1500
.ADD , , "Status", 1500
.ADD , , "Month", 0
End With
End Sub
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Sub loaddata()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from Attendence where EmpID like '" & EID & "%' and Month like '" & Month & "%' and Status like '" & Estatus & "%'", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!date)
list.SubItems(1) = rs!EmpID
list.SubItems(2) = rs!Name
list.SubItems(3) = rs!Status
list.SubItems(4) = rs!Month
rs.MoveNext
Loop
rs.Close
connect.Close
End Sub

Sub loaddata1()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from Attendence where EmpID like '" & EID & "%' and Status like '" & Estatus & "%'", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!date)
list.SubItems(1) = rs!EmpID
list.SubItems(2) = rs!Name
list.SubItems(3) = rs!Status
list.SubItems(4) = rs!Month
rs.MoveNext
Loop
rs.Close
connect.Close
End Sub



