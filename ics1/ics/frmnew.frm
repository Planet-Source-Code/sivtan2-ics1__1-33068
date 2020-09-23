VERSION 5.00
Begin VB.Form frmnew 
   Caption         =   "New User Form"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      DataField       =   "User Name"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtCpassword 
      DataField       =   "Password"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtLinkPwd 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      DataField       =   "Pass"
      DataSource      =   "data2"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Data data2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\VBPROJECT\ics\Login.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Login"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
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
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbUserName 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "MemID"
      DataSource      =   "data2"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Confirm Password :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Name :"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
frmmain.Show
Unload frmnew
End Sub
Private Sub cmd_Cancel_Click()
 txtUserName = ""
    txtCpassword = ""
    txtPassword = ""
    
        frmnew.Hide
        frmmain.Show
End Sub

Private Sub cmdOK_Click()
  Dim OnOff As Boolean

          If txtCpassword.Text <> txtPassword.Text Then
            MsgBox "Invalid Password"
            OnOff = False
             txtCpassword.SetFocus
             SendKeys "{Home}+{End}"
        Else
           data2.Recordset.AddNew
           lbUserName = txtUserName
           txtLinkPwd = txtPassword
           data2.Recordset.Update
       
           MsgBox " Password approved! ", vbOKOnly
                
           
            txtUserName = ""
            txtCpassword = ""
            txtPassword = ""
            Unload Me
            frmmain.Show
            
            Exit Sub
        End If
End Sub


