VERSION 5.00
Begin VB.Form frmnewuser 
   Caption         =   "New User"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\VBPROJECT\ics\pssw.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "id"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtcpssw 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   " "
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtpssw 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   " "
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtun 
      DataField       =   " "
      DataSource      =   "Data6"
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   " "
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdpssw 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmnewuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rec As Recordset
Private Sub Command2_Click()
Unload frmnewuser
frmchgpssw.Show
End Sub

Private Sub cmdcancel_Click()
Unload frmnewuser
frmchgpssw.Show
End Sub



Private Sub cmdpssw_Click()
   Dim OnOff As Boolean

          If txtcpssw.Text <> txtpssw.Text Then
            MsgBox "Invalid Password"
            OnOff = False
             txtcpssw.SetFocus
             SendKeys "{Home}+{End}"
        Else
           Data6.Recordset.AddNew
           lbUserName = txtun
           txtLinkPwd = txtpssw
           Data6.Recordset.Update
       
           MsgBox " Password approved! ", vbOKOnly
                
           
            txtun = ""
            txtcpssw = ""
            txtpssw = ""
            Unload Me
            frmmain.Show
            
            Exit Sub
        End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\pssw.mdb")
Set rec = db.OpenRecordset("select * from id")

End Sub
