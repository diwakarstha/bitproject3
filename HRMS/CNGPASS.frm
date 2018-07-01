VERSION 5.00
Begin VB.Form CNGPASS 
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CCANCEL 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CCONFIRM 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox OPASSTXT 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox NPASSTXT 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox CPASSTXT 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label CPASS 
      Caption         =   "Confirm Password :"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label NPASS 
      Caption         =   "New Password :"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label OPASS 
      Caption         =   "Old Password :"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "CNGPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub

