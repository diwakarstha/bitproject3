VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form EDETAIL 
   Caption         =   "Human Resourse Management System"
   ClientHeight    =   11835
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   21630
   LinkTopic       =   "Form2"
   Picture         =   "EDETAIL.frx":0000
   ScaleHeight     =   11835
   ScaleWidth      =   21630
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   6000
      Left            =   1005
      TabIndex        =   20
      Top             =   2700
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   10583
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
   Begin VB.TextBox AIDTXT 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   18480
      TabIndex        =   11
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox ANAMETXT 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   18480
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox AEMAILTXT 
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
      Left            =   18480
      TabIndex        =   9
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox Text5 
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
      Left            =   18480
      MaxLength       =   10
      TabIndex        =   8
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   18480
      TabIndex        =   7
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox Text3 
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
      Left            =   18480
      TabIndex        =   6
      Top             =   7200
      Width           =   2775
   End
   Begin VB.TextBox Text4 
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
      Left            =   18480
      MaxLength       =   10
      TabIndex        =   5
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
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
      Left            =   2880
      TabIndex        =   3
      Text            =   "Enter EmpID"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton EDELETE 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13800
      TabIndex        =   2
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton EDSA 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit / Update"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   1
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton ADD 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   0
      Top             =   9240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   18480
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Format          =   109707265
      CurrentDate     =   43245
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label AEMAIL 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   19
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label ANAME 
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
      Height          =   375
      Left            =   16560
      TabIndex        =   18
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label EID 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp. ID :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "D.O.B :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No. :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Post :"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Height          =   495
      Left            =   16560
      TabIndex        =   12
      Top             =   8040
      Width           =   1335
   End
End
Attribute VB_Name = "EDETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim selected As Boolean
Dim editmode As Boolean


Private Sub ADD_Click()
If editmode = False Then
    
    If ADD.Caption = "Add" Then
    If ListView1.ListItems.Count = 0 Then
    AIDTXT.Text = 1
    Else
    ListView1.ListItems(ListView1.ListItems.Count).selected = True
    AIDTXT.Text = ListView1.SelectedItem + 1
    End If
    ANAMETXT.Enabled = True
    DTPicker1.Enabled = True
    Text5.Enabled = True
    AEMAILTXT.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    ANAMETXT.Text = Empty
    Text5.Text = Empty
    AEMAILTXT.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty

    ADD.Caption = "Save"
    EDSA.Caption = "Cancel"
    EDELETE.Enabled = False
    ListView1.Enabled = False
    Text1.Enabled = False
    selected = True
    
    Else
    If Not (Text5.ForeColor = vbRed Or AEMAILTXT.ForeColor = vbRed Or ANAMETXT.Text = Empty Or Text5.Text = Empty Or AEMAILTXT.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty) Then
    rs.Open "select * from employee", connect, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields("EmpID") = AIDTXT.Text
    rs.Fields("Name") = ANAMETXT.Text
    rs.Fields("DOB") = DTPicker1.Value
    rs.Fields("Phone") = Text5.Text
    rs.Fields("Email") = AEMAILTXT.Text
    rs.Fields("Address") = Text2.Text
    rs.Fields("Post") = Text3.Text
    rs.Fields("Salary") = Text4.Text
    rs.Update
    ListView1.ListItems.Clear
    connect.Close
    loaddata
    ADD.Caption = "Add"
    EDSA.Caption = "Edit / Update"
    EDELETE.Enabled = True
    ListView1.Enabled = True
    AIDTXT.Enabled = False
    ANAMETXT.Enabled = False
    DTPicker1.Enabled = False
    Text5.Enabled = False
    AEMAILTXT.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text1.Enabled = True
    Else
    MsgBox "Please Fill all the Fields or check if all fields are valid"
    End If
    End If
Else
If Not (Text5.ForeColor = vbRed Or AEMAILTXT.ForeColor = vbRed Or ANAMETXT.Text = Empty Or Text5.Text = Empty Or AEMAILTXT.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty) Then
Y = ListView1.SelectedItem
rs.Open "select * from employee where EmpID=" & Y, connect, adOpenDynamic, adLockOptimistic
'rs.Fields("EmpID") = AIDTXT.Text
rs.Fields("Name") = ANAMETXT.Text
rs.Fields("DOB") = DTPicker1.Value
rs.Fields("Phone") = Text5.Text
rs.Fields("Email") = AEMAILTXT.Text
rs.Fields("Address") = Text2.Text
rs.Fields("Post") = Text3.Text
rs.Fields("Salary") = Text4.Text
rs.Update
connect.Close
loaddata
ADD.Caption = "Add"
EDSA.Caption = "Edit / Update"
EDELETE.Enabled = True
ListView1.Enabled = True
AIDTXT.Enabled = False
ANAMETXT.Enabled = False
DTPicker1.Enabled = False
Text5.Enabled = False
AEMAILTXT.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text1.Enabled = True
editmode = False
'MsgBox "FOR EDIT MODE", , "Message"
Else
    MsgBox "Please Fill all the Fields or check if all fields are valid"
    End If
End If
End Sub

Private Sub AEMAILTXT_LostFocus()
If Not AEMAILTXT.Text = Empty Then
If IsValidEmail(AEMAILTXT.Text) = False Then
MsgBox "Please Enter Valid Email!!"
AEMAILTXT.ForeColor = vbRed
Else
AEMAILTXT.ForeColor = vbBlack
End If
End If
End Sub

Private Sub ANAMETXT_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub EDELETE_Click()
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then

rs.Open "select * from employee where EmpID =" & ListView1.SelectedItem, connect, adOpenDynamic, adLockOptimistic
rs.Delete
connect.Close
loaddata
Else
MsgBox "Record Not Deleted!", , "Message"
End If
End Sub


Private Sub EDSA_Click()
If selected = True Then
If EDSA.Caption = "Edit / Update" Then
AIDTXT.Enabled = True
ANAMETXT.Enabled = True
DTPicker1.Enabled = True
Text5.Enabled = True
AEMAILTXT.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text1.Enabled = False
ADD.Caption = "Save"
EDSA.Caption = "Cancel"
EDELETE.Enabled = False
ListView1.Enabled = False
editmode = True
Else
AIDTXT.Text = " "
ANAMETXT.Text = " "
Text5.Text = " "
AEMAILTXT.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
ADD.Caption = "Add"
EDSA.Caption = "Edit / Update"
Text1.Enabled = True
EDELETE.Enabled = True
ListView1.Enabled = True
AIDTXT.Enabled = False
ANAMETXT.Enabled = False
DTPicker1.Enabled = False
Text5.Enabled = False
AEMAILTXT.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
selected = False
editmode = False
End If
Else
MsgBox "Record Not selected! Please select any one record.", , "Message"
End If
End Sub


Private Sub Form_Load()
Me.Picture = LoadPicture("Pictures\entry.jpg")
With ListView1.ColumnHeaders
.ADD , , "EmpID", 1500
.ADD , , "Name", 2000
.ADD , , "DOB", 1500
.ADD , , "Email", 2500
.ADD , , "Mobile No.", 2000
.ADD , , "Address", 2000
.ADD , , "Post", 1500
.ADD , , "Salary", 1600
End With
loaddata
AIDTXT.Enabled = False
ANAMETXT.Enabled = False
DTPicker1.Enabled = False
Text5.Enabled = False
AEMAILTXT.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
DTPicker1.MaxDate = Now
DTPicker1.Value = Now

End Sub
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Sub loaddata()
Dim list As ListItem
ListView1.ListItems.Clear
dbconnection
rs.Open "select * from employee", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!DOB
list.SubItems(3) = rs!email
list.SubItems(4) = rs!Phone
list.SubItems(5) = rs!Address
list.SubItems(6) = rs!Post
list.SubItems(7) = rs!Salary
rs.MoveNext
Loop
rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Show
connect.Close
End Sub

Private Sub ListView1_Click()
On Error Resume Next
selected = True
AIDTXT.Text = ListView1.SelectedItem
ANAMETXT.Text = ListView1.SelectedItem.SubItems(1)
DTPicker1.Value = ListView1.SelectedItem.SubItems(2)
AEMAILTXT.Text = ListView1.SelectedItem.SubItems(3)
Text5.Text = ListView1.SelectedItem.SubItems(4)
Text2.Text = ListView1.SelectedItem.SubItems(5)
Text3.Text = ListView1.SelectedItem.SubItems(6)
Text4.Text = ListView1.SelectedItem.SubItems(7)
End Sub

Private Sub Text1_Change()
ListView1.ListItems.Clear
rs.Open "select * from employee where EmpID like '" & Text1 & "%'", connect, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView1.ListItems.ADD(, , rs!EmpID)
list.SubItems(1) = rs!Name
list.SubItems(2) = rs!DOB
list.SubItems(3) = rs!email
list.SubItems(4) = rs!Phone
list.SubItems(5) = rs!Address
list.SubItems(6) = rs!Post
list.SubItems(7) = rs!Salary
rs.MoveNext
Loop
rs.Close
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub
Public Function IsValidEmail(email As String) As Boolean
Dim myAt As Integer
Dim myAtLastPos As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim myDotAt As Integer
Dim myAtDot As Integer
Dim mySpace As Integer
IsValidEmail = True
mySpace = InStr(1, email, " ", vbTextCompare)
myAtLastPos = InStrRev(email, "@", , vbTextCompare)
myAt = InStr(1, email, "@", vbTextCompare)
myAtDot = InStr(1, email, "@.", vbTextCompare)
myDotAt = InStr(1, email, ".@", vbTextCompare)
myDot = InStr(myAt + 2, email, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
If myAtDot > 0 Or myDotAt > 0 Or myAtLastPos <> myAt Or mySpace > 0 Or myAt = 0 Or myDot = 0 Or myDotDot > 0 Or Right(email, 1) = "." Then IsValidEmail = False
End Function

Private Sub Text5_LostFocus()
If Not Text5.Text = Empty Then
If Not Len(Text5.Text) = 10 Then
MsgBox "Please Enter Valid Mobile Number of 10 Digits!!"
Text5.ForeColor = vbRed
Else
Text5.ForeColor = vbBlack
End If
End If
End Sub
