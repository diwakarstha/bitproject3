VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EADD 
   Caption         =   "ADD EMPLOYEE"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Text            =   "Combo3"
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2880
      TabIndex        =   18
      Text            =   "Combo2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   1560
      Width           =   855
   End
   Begin MSAdodcLib.Adodc EmpDatabase 
      Height          =   615
      Left            =   5160
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\WELCOME\Desktop\HRMS\Database\Employee.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\WELCOME\Desktop\HRMS\Database\Employee.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
      Caption         =   "Employee Database"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text4 
      DataField       =   "Salary"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Post"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Address"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Phone"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox AEMAILTXT 
      DataField       =   "Email"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox ANAMETXT 
      DataField       =   "Name"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox AIDTXT 
      DataField       =   "EmpID"
      DataSource      =   "EmpDatabase"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Salary :"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Post :"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Address :"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Phone :"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "D.O.B :"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label EID 
      Caption         =   "Emp. ID :"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label ANAME 
      Caption         =   "Name :"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label AEMAIL 
      Caption         =   "Email :"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "EADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
EmpDatabase.Recordset.AddNew

End Sub

Private Sub Command2_Click()
EmpDatabase.Recordset.CancelBatch
Employee.Show
EADD.Hide
End Sub

