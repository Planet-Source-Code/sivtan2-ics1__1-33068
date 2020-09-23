VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Log In"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Pass"
      DataSource      =   "Data1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "MemID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\VBPROJECT\ics\Login.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Login"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Pass 
      Caption         =   " "
      DataField       =   "Pass"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label MemID 
      Caption         =   " "
      DataField       =   "MemID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Welcome"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ICS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
'Dim rec As Recordset


Private Sub cmd_ok_Click()
  Dim Chkpwd As Boolean
      Chkpwd = True
    If Len(Text3.Text) <> 0 And Len(Text4.Text) <> 0 Then
     Do
        If Text1.Text <> Text3.Text Then
            If Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                MsgBox "Invalid User!!"
                Text3.SetFocus
                SendKeys "{Home}+{End}"
                Text4.Text = ""
                Chkpwd = False
             
        Else
            Data1.Recordset.MoveNext
        End If
    Else 'if the username match
        If Text2.Text <> Text4.Text Then
         MsgBox "Invalid Password, try again!", , "Login"
            
            Text4.SetFocus
            SendKeys "{Home}+{End}"
            Chkpwd = False
            Data1.Recordset.MoveFirst
        Else
            Unload Login
            frmmain.Show
            MsgBox "Login Succeed!"
            Exit Sub
        End If
    End If
    Loop While Chkpwd <> False
Else
    MsgBox "You should enter user name and user Password!"
End If
End Sub

Private Sub cmdcancel_Click()
    'Unload Me
    End
End Sub

Private Sub Form_Load()
'Set db = OpenDatabase(App.Path & "\login.mdb")
'Set rec = db.OpenRecordset("select * from login")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmd_ok.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = Chr(KeyAscii) Like "[A-Za-z]" Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 13
'Text1.Locked = True
If eno = True Then
Text1.Locked = False
Text1.Visible = True
If KeyAscii = 13 Then
If Not (Text1.Text) = "" Then
Text2.SetFocus
Else
MsgBox "THE USER NAME IS EMPTY"
Text1.SetFocus
Text1.Text = ""
Text2.Locked = True
End If
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = Chr(KeyAscii) Like "[A-Za-z]" Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 13
'Text2.Locked = True
If eno = True Then
Text2.Locked = False
If KeyAscii = 13 Then
If Not (Text2.Text) = "" Then
cmd_ok.SetFocus
Else
MsgBox "THE PASSWORD IS EMPTY"
Text2.SetFocus
Text2.Text = ""
End If
End If
End If
End Sub
