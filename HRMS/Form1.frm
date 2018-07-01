VERSION 5.00
Begin VB.Form Mainmenu 
   Caption         =   "Human Resourse Management System"
   ClientHeight    =   5670
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton MCREPORT 
      Caption         =   "Company Report"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton MPROJECT 
      Caption         =   "Project"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton MEMPLOYEE 
      Caption         =   "Employee"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Menu ADMIN 
      Caption         =   "ADMIN"
      Begin VB.Menu cp 
         Caption         =   "change password"
      End
      Begin VB.Menu ai 
         Caption         =   "Admin info"
      End
      Begin VB.Menu lo 
         Caption         =   "Log Out"
      End
   End
   Begin VB.Menu helpdesk 
      Caption         =   "HELP"
   End
End
Attribute VB_Name = "Mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MEMPLOYEE_Click()
Employee.Show
EDETAIL.Show
End Sub
