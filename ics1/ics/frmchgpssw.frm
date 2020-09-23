VERSION 5.00
Begin VB.Form frmchgpssw 
   Caption         =   "Change Password"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdcfm 
      Caption         =   "Confirm"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New User"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtpssw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtun 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5175
      Begin VB.TextBox txtcpssw 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtnew 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Password"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "New Password"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
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
      Left            =   4440
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmchgpssw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rec As Recordset

Private Sub Command4_Click()
Unload frmchgpssw
frmmain.Show
End Sub
Public Sub editpwd(ans)
    If ans = 1 Then
         txtCpassword.Visible = True
   
   
    Else
        txtCpassword.Visible = False
    
    End If
End Sub

Private Sub cmdcancel_Click()
Unload frmchgpssw
frmmain.Show
End Sub

Private Sub cmdcfm_Click()
 Dim OnOff As Boolean
        OnOff = True
      Do
          If txtnew.Text <> txtcpssw.Text Then
            MsgBox "Invalid Password"
                txtcpssw.SetFocus
             SendKeys "{Home}+{End}"
          
            OnOff = False
        Else
            rec.Edit
            
            rec.Fields(2).Value = txtcpssw.Text
            rec.Update
            MsgBox " Your Password had been Changed! ", vbOKOnly
                       
            txtnew.Enabled = False
            txtcpssw.Enabled = False
            Label3.Enabled = False
            Label4.Enabled = False
                
            cmdcfm.Enabled = False
            cmdok.Enabled = True
          
            txtpssw.Text = ""
            txtun.Text = ""
            txtcpssw.Text = ""
            txtnew.Text = ""
            Unload Me
            frmmain.Show
            Exit Sub
        End If
    
    Loop While OnOff <> False
End Sub

Private Sub cmdnew_Click()
frmnewuser.Show
frmchgpssw.Hide
End Sub
Private Sub cmdok_Click()
 Set db = OpenDatabase(App.Path & "\pssw.mdb")
Set rec = db.OpenRecordset("select * from id")
 Dim Chkpwd As Boolean
      Chkpwd = True
    If Len(txtun.Text) <> 0 And Len(txtpssw.Text) <> 0 Then
    Do
    If rec.Fields(2).Value <> txtun.Text Then
        If rec.EOF Then
            rec.MoveFirst
            MsgBox "Invalid User!!"
            txtun.SetFocus
            Chkpwd = False
            SendKeys "{Home}+{End}"
        Else
            rec.MoveNext
           
        End If
        
    Else 'if the username match
        If rec.Fields(2).Value <> (txtpssw.Text) Then
         MsgBox "Invalid Password, try again!", , "Login"
            
            txtpssw.SetFocus
            SendKeys "{Home}+{End}"
            Chkpwd = False
           
        Else
            
            txtnew.Enabled = True
           
            Label3.Enabled = True
            Label4.Enabled = True
            txtpssw.Enabled = True
            cmdcfm.Enabled = True
            cmdok.Enabled = False
            txtnew.SetFocus
            SendKeys "{Home}+{End}"
            txtpssw.Text = ""
            Exit Sub
        End If
     End If
     
    Loop While Chkpwd <> False
Else
    MsgBox "You should enter user name and user Password!"
             txtun.SetFocus
             SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\pssw.mdb")
Set rec = db.OpenRecordset("select * from id")
'Data5.DatabaseName = WorkPath + PwdFile
'cmdok.Enabled = False
End Sub

