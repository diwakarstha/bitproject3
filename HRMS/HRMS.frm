VERSION 5.00
Begin VB.Form HRMS 
   Caption         =   "Employee Record Management System"
   ClientHeight    =   12375
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17040
   LinkTopic       =   "Form1"
   Picture         =   "HRMS.frx":0000
   ScaleHeight     =   21.828
   ScaleMode       =   0  'User
   ScaleWidth      =   30.057
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame ADL 
      BackColor       =   &H00808080&
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   8640
      TabIndex        =   0
      Top             =   4200
      Width           =   6135
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox PASSTXT 
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
         IMEMode         =   3  'DISABLE
         Left            =   2520
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
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
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton ALOGIN 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   1
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label PASS 
         BackColor       =   &H00808080&
         Caption         =   "Password :"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label AID 
         BackColor       =   &H00808080&
         Caption         =   "Admin ID :"
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
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "HRMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub AEXIT_Click()
End
End Sub
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Private Sub ALOGIN_Click()
dbconnection
rs.Open "select * from Admin", connect, adOpenDynamic, adLockOptimistic
If rs.Fields("ID") = AIDTXT.Text And rs.Fields("Password") = PASSTXT.Text Then
connect.Close
Menu.Show
'EDETAIL.Show
HRMS.Hide
Else
MsgBox "Please check AdminID or password is wrong!!"
connect.Close
End If
End Sub

Private Sub Exit_Click()
End
End Sub


Private Sub Form_Load()
Me.Picture = LoadPicture("Pictures\PIC1.jpg")
AIDTXT.Enabled = False
AIDTXT.Text = "1527812018"
End Sub
