VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmview 
   Caption         =   "View Balance"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form4"
   ScaleHeight     =   4590
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\My Documents\VBPROJECT\ics\db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\My Documents\VBPROJECT\ics\db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from ics"
      Caption         =   "View Balance"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Main"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4200
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmview.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   3120
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Stock Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ICS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rec As Recordset
Private Sub Command1_Click()
Unload frmview
frmmain.Show
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("c:\mydocu~1\VBPROJECT\ics\db1.mdb")
Set rec = db.OpenRecordset("select * from ics")

Label3.Caption = Date
Label4.Caption = Time
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub
