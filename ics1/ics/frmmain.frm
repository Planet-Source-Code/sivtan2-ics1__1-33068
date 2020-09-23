VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "ICS-MAIN MENU "
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "New User"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add New Record"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmd_view 
      Caption         =   "&View Balance"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmd_stockout 
      Caption         =   "Stock &Out"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmd_stockin 
      Caption         =   "Stock &In"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   600
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   2880
   End
   Begin VB.Label Label5 
      Caption         =   "2002  by tanhuathiam"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label label 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Copyright "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INVENTORY CONTROL SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_chgpssw_Click()
frmchgpssw.Show
frmmain.Hide
End Sub

Private Sub cmd_exit_Click()
msg = MsgBox("Are you sure want to exit ??", vbOKCancel, "Exit")
If msg = vbOK Then
End
Else
frmmain.Show
End If
End Sub

Private Sub cmd_stockin_Click()
frmstkin.Show
frmmain.Hide
End Sub

Private Sub cmd_stockout_Click()
frmstkout.Show
frmmain.Hide
End Sub

Private Sub cmd_view_Click()
frmview.Show
frmmain.Hide
End Sub

Private Sub cmdadd_Click()
frmadd.Show
frmmain.Hide
End Sub

Private Sub cmdsearch_Click()
frmsearch.Show
frmmain.Hide
End Sub

Private Sub Command1_Click()
frmnew.Show
frmmain.Hide

End Sub

Private Sub Form_Load()
Login.Show
frmmain.Hide
Label2.Caption = Date
Label3.Caption = Time
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Time
End Sub



