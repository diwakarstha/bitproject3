VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Form1"
   ClientHeight    =   10845
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   21660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Mainmenu.frx":0000
   ScaleHeight     =   10000
   ScaleMode       =   0  'User
   ScaleWidth      =   27583.25
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Record Management System"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Image menubg 
      Height          =   3615
      Left            =   1440
      Top             =   5760
      Width           =   5055
   End
   Begin VB.Label LO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18600
      TabIndex        =   5
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Label CAI 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Change Admin   Info"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13080
      TabIndex        =   4
      Top             =   9000
      Width           =   3015
   End
   Begin VB.Label ER 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Employee Report"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   3
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Label MS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Manage   Salary"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18480
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label MA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Attendance"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13440
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label MEmp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Employee"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   0
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Image EPIC 
      Height          =   2175
      Left            =   8160
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Image LOPIC 
      Height          =   2175
      Left            =   18480
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Image CAPIC 
      Height          =   2175
      Left            =   13320
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Image ERPIC 
      Height          =   2175
      Left            =   8280
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Image SPIC 
      Height          =   2175
      Left            =   18360
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Image ATPIC 
      Height          =   2175
      Left            =   13320
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim over As Boolean
Private Sub ATPIC_Click()
Attendence.Show
Menu.Hide
End Sub

Private Sub CAI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CAI.FontSize = 18
CAI.ForeColor = vbBlue
over = True
End Sub

Private Sub CAPIC_Click()
AINFO.Show
Menu.Hide
End Sub

Private Sub CAPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CAI.FontSize = 18
CAI.ForeColor = vbBlue
over = True
End Sub

Private Sub EPIC_Click()
EDETAIL.Show
Menu.Hide
End Sub

Private Sub EPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MEmp.FontSize = 18
MEmp.ForeColor = vbBlue
over = True
End Sub

Private Sub ER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ER.FontSize = 18
ER.ForeColor = vbBlue
over = True
End Sub

Private Sub ERPIC_Click()
Empreport.Show
Menu.Hide
End Sub

Private Sub ERPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ER.FontSize = 18
ER.ForeColor = vbBlue
over = True
End Sub

Private Sub Form_Load()
EPIC.Picture = LoadPicture("Pictures\emp.gif")
ATPIC.Picture = LoadPicture("Pictures\Attend.gif")
SPIC.Picture = LoadPicture("Pictures\salary.gif")
ERPIC.Picture = LoadPicture("Pictures\erep.gif")
CAPIC.Picture = LoadPicture("Pictures\cai.gif")
LOPIC.Picture = LoadPicture("Pictures\lo.gif")
menubg.Picture = LoadPicture("Pictures\menubg.gif")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If over = True Then
MEmp.FontSize = 14
MEmp.ForeColor = vbBlack
MA.FontSize = 14
MA.ForeColor = vbBlack
MS.FontSize = 14
MS.ForeColor = vbBlack
ER.FontSize = 14
ER.ForeColor = vbBlack
CAI.FontSize = 14
CAI.ForeColor = vbBlack
LO.FontSize = 14
LO.ForeColor = vbBlack
over = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub LO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LO.FontSize = 18
LO.ForeColor = vbBlue
over = True
End Sub

Private Sub LOPIC_Click()
Menu.Hide
HRMS.PASSTXT.Text = Empty
HRMS.Show
End Sub

Private Sub LOPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LO.FontSize = 18
LO.ForeColor = vbBlue
over = True
End Sub

Private Sub ATPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MA.FontSize = 18
MA.ForeColor = vbBlue
over = True
End Sub

Private Sub MA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MA.FontSize = 18
MA.ForeColor = vbBlue
over = True
End Sub

Private Sub MEmp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MEmp.FontSize = 18
MEmp.ForeColor = vbBlue
over = True
End Sub

Private Sub MS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MS.FontSize = 18
MS.ForeColor = vbBlue
over = True
End Sub

Private Sub SPIC_Click()
Salary.Show
Menu.Hide
End Sub

Private Sub SPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MS.FontSize = 18
MS.ForeColor = vbBlue
over = True
End Sub
