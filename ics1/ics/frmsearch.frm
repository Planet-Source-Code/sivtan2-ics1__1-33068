VERSION 5.00
Begin VB.Form frmstkin 
   Caption         =   "Stock In"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdview 
      Caption         =   "View Balance"
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Main"
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdstockin 
      Caption         =   "Stock In"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtproduct 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtmc 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtcp 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   " "
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtrp 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtpd 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   " "
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtqty 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " "
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3600
      Top             =   4680
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
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Model Code"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Cost Price"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Retail Pricre"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Production Date"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Stock In Form"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmstkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rec As Recordset

Private Sub cmdcancel_Click()
txtproduct.Text = ""
txtcp.Text = ""
txtrp.Text = ""
txtpd.Text = ""
txtmc.Text = ""
txtqty.Text = ""
cmdsearch.Enabled = True
cmdstockin.Enabled = False
End Sub

Private Sub cmdmain_Click()
frmmain.Show
Unload frmstkin
End Sub

Private Sub cmdsearch_Click()
Set db = OpenDatabase(App.Path & "\db1.mdb")
Set rec = db.OpenRecordset("select * from ics where product = '" & txtproduct.Text & "'")
If Not (txtproduct.Text) = "" Then

 If rec.RecordCount = 0 Then
 MsgBox "record does not exist !!!!!", vbCritical
 Else
 txtcp.Text = rec.Fields(2)
 txtrp.Text = rec.Fields(3)
 txtpd.Text = rec.Fields(4)
 txtmc.Text = rec.Fields(1)
 txtqty.Text = rec.Fields(5)

 cmdstockin.Enabled = True
 cmdsearch.Enabled = False
 End If
Else
 MsgBox "You never enter the product!", vbInformation, "Error"
 txtproduct.SetFocus
End If

End Sub

Private Sub cmdstockin_Click()
Dim num As Variant
Dim sum As Variant

num = InputBox("please enter the quantity you want to store in", "Stock In", "0")


If num <> "" Then

sum = Val(txtqty.Text) + Val(num)
MsgBox "YOUR QUANTITY NOW IS " & sum
txtqty.Text = sum
rec.Edit
rec.Fields(5) = sum
rec.Update
cmdstockin.Enabled = False
cmdsearch.Enabled = True
Else
MsgBox "Sorry!! Nothing to stock in. Check Value that you're entered"
cmdsearch.Enabled = True
cmdstockin.Enabled = False
End If

End Sub


Private Sub cmdview_Click()
frmview.Show
frmstkin.Hide
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\db1.mdb")
Set rec = db.OpenRecordset("select * from ics where product = '" & txtproduct.Text & "'")
Label9.Caption = Date
Label10.Caption = Time

End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub

