VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Log In"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
