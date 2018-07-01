VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form empat 
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17610
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   17610
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   4320
      TabIndex        =   0
      Top             =   2160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12303
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
End
Attribute VB_Name = "empat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
With ListView1.ColumnHeaders
.ADD , , "EmpID", 1500
.ADD , , "Name", 1500
.ADD , , "Phone", 1500
End With
loaddata
End Sub
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub
Sub loaddata()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from Employee", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!Phone
rs.MoveNext
Loop
rs.Close
connect.Close
End Sub

