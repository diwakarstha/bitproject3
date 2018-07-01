VERSION 5.00
Begin VB.MDIForm Employee 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   11700
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu ED 
      Caption         =   "Employee Detail"
   End
   Begin VB.Menu A 
      Caption         =   "Attendence"
   End
   Begin VB.Menu S 
      Caption         =   "Salary"
   End
   Begin VB.Menu Ad 
      Caption         =   "Admin"
      Begin VB.Menu CP 
         Caption         =   "Change Password"
      End
      Begin VB.Menu UAI 
         Caption         =   "Update Admin Info"
      End
      Begin VB.Menu LO 
         Caption         =   "Log Out"
      End
   End
End
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_Click()
Attendence.Show
EDETAIL.Hide
Salary.Hide
End Sub

Private Sub CP_Click()
CNGPASS.Show
End Sub

Private Sub ED_Click()
EDETAIL.Show
Attendence.Hide
Salary.Hide
End Sub



Private Sub LO_Click()
HRMS.Show
Employee.Hide
End Sub



Private Sub MDIForm_Load()
'EPIC.Picture = LoadPicture("D:\HRMS\Pictures\emp.gif")
End Sub

Private Sub S_Click()
Salary.Show
EDETAIL.Hide
Attendence.Hide
End Sub

Private Sub UAI_Click()
AINFO.Show
End Sub
