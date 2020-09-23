VERSION 5.00
Begin VB.Form frmstkout 
   Caption         =   "Stock Out"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdview 
      Caption         =   "View Balance"
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Main"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdstkout 
      Caption         =   "Stock Out"
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtproduct 
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtqty 
      DataField       =   " "
      DataSource      =   " "
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   " "
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtpd 
      DataField       =   " "
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   " "
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtrp 
      DataField       =   " "
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtcp 
      DataField       =   " "
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   " "
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtmc 
      DataField       =   " "
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   4680
   End
   Begin VB.Label Label10 
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Stock Out"
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
      Left            =   1560
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Production Date"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Retail Pricre"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Cost Price"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Model Code"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmstkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
cmdstkout.Enabled = False
End Sub

Private Sub cmdmain_Click()
Unload frmstkout
frmmain.Show
End Sub

Private Sub cmdsearch_Click()
Set db = OpenDatabase(App.Path & "\db1.mdb")
Set rec = db.OpenRecordset("select * from ics where product = '" & txtproduct.Text & "'")
If Not (txtproduct.Text) = "" Then

 If rec.RecordCount = 0 Then
 MsgBox "Sorry !! Record does not exist !!!!!", vbCritical
 Else
 txtcp.Text = rec.Fields(2)
 txtrp.Text = rec.Fields(3)
 txtpd.Text = rec.Fields(4)
 txtmc.Text = rec.Fields(1)
 txtqty.Text = rec.Fields(5)
 cmdstkout.Enabled = True
 cmdsearch.Enabled = False
 End If

Else
 MsgBox "You never enter the product!", vbInformation, "Error"
 txtproduct.SetFocus
End If
End Sub

Private Sub cmdstkout_Click()
Dim num As Variant
Dim sum As Variant

num = InputBox("Please enter the quantity you want to store in ", "Stock Out", "0")
If num < Val(txtqty.Text) Then
rec.Edit
sum = Val(txtqty.Text) - Val(num)
MsgBox "YOUR QUANTITY NOW IS " & sum
txtqty.Text = sum
rec.Edit
rec.Fields(5) = sum
rec.Update
cmdstkout.Enabled = False
cmdsearch.Enabled = True
Else
MsgBox "Sorry!! Nothing to stock in. Check Value that you're entered"
cmdsearch.Enabled = True
cmdstockin.Enabled = False
End If
End Sub

Private Sub cmdview_Click()
frmview.Show
frmstkout.Hide
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\db1.mdb")
Set rec = db.OpenRecordset("select * from ics")
cmdstkout.Enabled = False
Label9.Caption = Date
Label10.Caption = Time
End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub
