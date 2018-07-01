VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ProjectPerf 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11400
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6840
      Top             =   10080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton EUPDATE 
      Caption         =   "View Detail"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton EREPORT 
      Caption         =   "Project Performance Report"
      Height          =   735
      Left            =   10080
      TabIndex        =   3
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton EFIND 
      Caption         =   "Find"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox ETABLE 
      Height          =   6000
      Left            =   1005
      ScaleHeight     =   5940
      ScaleWidth      =   15090
      TabIndex        =   1
      Top             =   2340
      Width           =   15150
   End
End
Attribute VB_Name = "ProjectPerf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = Employee.Width - 200
Me.Height = Employee.Height - 850
Me.Left = Employee.Left - 100
Me.Top = Employee.Top - 750
End Sub
