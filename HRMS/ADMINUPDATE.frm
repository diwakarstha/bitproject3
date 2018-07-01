VERSION 5.00
Begin VB.Form AINFO 
   Caption         =   "ADMIN INFO"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   Picture         =   "ADMINUPDATE.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   16755
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Change Admin Info"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   13935
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   1680
         TabIndex        =   7
         Top             =   2520
         Width           =   10695
         Begin VB.TextBox OPASSTXT 
            BeginProperty Font 
               Name            =   "Copperplate Gothic Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   13
            Top             =   1560
            Width           =   3735
         End
         Begin VB.TextBox NPASSTXT 
            BeginProperty Font 
               Name            =   "Copperplate Gothic Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   12
            Top             =   2520
            Width           =   3735
         End
         Begin VB.TextBox CPASSTXT 
            BeginProperty Font 
               Name            =   "Copperplate Gothic Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   11
            Top             =   3480
            Width           =   3735
         End
         Begin VB.Label CPASS 
            Caption         =   "Confirm Password :"
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
            Left            =   720
            TabIndex        =   16
            Top             =   3600
            Width           =   2535
         End
         Begin VB.Label NPASS 
            Caption         =   "New Password :"
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
            Left            =   720
            TabIndex        =   15
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label OPASS 
            Caption         =   "Old Password :"
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
            Left            =   720
            TabIndex        =   14
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "Algerian"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   10
            Top             =   3600
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Copperplate Gothic Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "Algerian"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7200
            TabIndex        =   8
            Top             =   1680
            Width           =   3135
         End
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
         Height          =   615
         Left            =   5040
         TabIndex        =   4
         Top             =   1680
         Width           =   3735
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
         Height          =   615
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7440
         TabIndex        =   2
         Top             =   7920
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         TabIndex        =   1
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "AINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Sub dbconnection()
connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Employee.mdb;Persist Security Info=False"
End Sub

Private Sub Command2_Click()
AINFO.Hide
Menu.Show
End Sub

Private Sub Command3_Click()
If Not Text2.Text = Empty Then
dbconnection
rs.Open "select * from Admin where ID=1527812018 ", connect, adOpenDynamic, adLockOptimistic
rs.Fields("Name") = Text2.Text
If Not Label3.Caption = "Cancel" Then
rs.Update
MsgBox "Name Updated succesfully"
connect.Close
End If
If Label3.Caption = "Cancel" Then

If Not (Label4.Caption = "Wrong Old Password" Or Label5.Caption = "Must Fill New Password!!" Or Label5.Caption = "New Password didnot Matched") Then
'dbconnection
'rs.Open "select * from Admin where ID=1527812018 ", connect, adOpenDynamic, adLockOptimistic
rs.Fields("Password") = CPASSTXT.Text
rs.Update
MsgBox "Name and Password Updated succesfully"
Else
MsgBox "Please check old password is correct or new password is matching"
End If
connect.Close
End If
Else
MsgBox "Must fill All the Fields"
End If
End Sub

Private Sub CPASSTXT_LostFocus()
If Not NPASSTXT.Text = Empty Then

If Not CPASSTXT.Text = Empty Then
If NPASSTXT.Text = CPASSTXT.Text Then
CPASSTXT.ForeColor = vbBlack
Label5.Caption = "New Password Matched"
Else
CPASSTXT.ForeColor = vbRed
Label5.Caption = "New Password didnot Matched"
End If
End If
Else
CPASSTXT.ForeColor = vbRed
Label5.Caption = "Must Fill New Password!!"
End If
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture("Pictures\entry.jpg")
dbconnection
rs.Open "select * from Admin where ID=1527812018 ", connect, adOpenDynamic, adLockOptimistic
Text1.Text = rs!ID
connect.Close
Text1.Enabled = False
Label3.ForeColor = vbBlue
OPASSTXT.Enabled = False
NPASSTXT.Enabled = False
CPASSTXT.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
AINFO.Hide
Menu.Show
End Sub

Private Sub Label3_Click()
If Label3.Caption = "Edit" Then
Label3.ForeColor = vbRed
Label3.Caption = "Cancel"
OPASSTXT.Enabled = True
NPASSTXT.Enabled = True
CPASSTXT.Enabled = True
Else
Label3.ForeColor = vbBlue
Label3.Caption = "Edit"
OPASSTXT.Enabled = False
NPASSTXT.Enabled = False
CPASSTXT.Enabled = False
OPASSTXT.Text = Empty
NPASSTXT.Text = Empty
CPASSTXT.Text = Empty
End If
End Sub

Private Sub OPASSTXT_LostFocus()
If Not OPASSTXT.Text = Empty Then
dbconnection
rs.Open "select * from Admin where ID=1527812018 ", connect, adOpenDynamic, adLockOptimistic
If rs.Fields("Password") = OPASSTXT.Text Then
OPASSTXT.ForeColor = vbBlack
Label4.Caption = "Correct Old Password"
connect.Close
Else
OPASSTXT.ForeColor = vbRed
Label4.Caption = "Wrong Old Password"
connect.Close
End If
End If
End Sub
