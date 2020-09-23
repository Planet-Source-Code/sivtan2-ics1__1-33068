VERSION 5.00
Begin VB.Form frmstkin 
   Caption         =   "Stock In"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form6"
   ScaleHeight     =   5265
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstModel 
      Height          =   1425
      Left            =   4800
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   4680
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\VBPROJECT\ics\db1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ics"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbproduct 
      Height          =   315
      ItemData        =   "frmstkin.frx":0000
      Left            =   1800
      List            =   "frmstkin.frx":0002
      TabIndex        =   14
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtqty 
      DataField       =   "quantity"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   " "
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtpd 
      DataField       =   "date"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Text            =   " "
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtrp 
      DataField       =   "reprice"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtcp 
      DataField       =   "cost"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   " "
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtmc 
      DataField       =   "model"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   " "
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Stock In"
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
      Left            =   1680
      TabIndex        =   15
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Production Date"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Retail Pricre"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Cost Price"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Model Code"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmstkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rec As Recordset

Private Sub cmbproduct_Click()
LoadModel cmbproduct.Text
End Sub

Private Sub cmdcancel_Click()
Unload frmstkin
frmmain.Show
End Sub

Sub LoadDetail(mdlName As String)
Dim CurrBkMrk As String
   Dim fldName As String
   On Error GoTo LoadErr
    If mdlName = "" Then Exit Sub
    'lstModel.Clear
    currbkmark = Data2.Recordset.Bookmark
    Data2.Recordset.MoveFirst
    
    Do Until (Data2.Recordset.EOF)
        fldName = Data2.Recordset.Fields(1).Value
        If mdlName = fldName Then
            txtcp = Data2.Recordset.Fields(2).Value
            txtmn = Data2.Recordset.Fields(3).Value
            txtmc = Data2.Recordset.Fields(4).Value
            currbkmark = Data2.Recordset.Bookmark
            Data2.Recordset.MoveLast
        End If
        Data2.Recordset.MoveNext
    Loop
    
    Data2.Recordset.Bookmark = currbkmark
LoadErr:
    Resume Next
'Next
End Sub

Private Sub cmdok_Click()
Dim response As Integer
        response = MsgBox("Are you sure ?", vbYesNo + vbQuestion, "Stock In")
    If response = vbYes Then
    
        lbStockQty = Val(lbStockQty) + Val(txtQuantity.Text)
        txtqty.Text = ""
  
 Else
        txtqty.Text = ""
    Exit Sub
 End If
End Sub

Private Sub Command1_Click()
 LoadCmb
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("c:\mydocu~1\VBPROJECT\ics\db1.mdb")
Set rec = db.OpenRecordset("select * from ics")


Label9.Caption = Date
Label10.Caption = Time
End Sub

Private Sub List1_DblClick()
LoadDetail Mid(lstModel.Text, 3)

End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub
Sub LoadCmb()
   Dim CurrBkMrk As String
   Dim fldName As String
    
    currbkmark = Data2.Recordset.Bookmark
    Data2.Recordset.MoveFirst
    
    Do Until (Data2.Recordset.EOF)
        fldName = Data2.Recordset.Fields(0).Value
        For i = 0 To cmbproduct.ListCount - 1
            cmbproduct.ListIndex = i
            If fldName = cmbproduct.Text Then
               i = cmbproduct.ListCount
               fldName = ""
            End If
        Next i
        If fldName <> "" Then
            cmbproduct.AddItem fldName
        End If
        Data2.Recordset.MoveNext
    Loop
    
    Data2.Recordset.Bookmark = currbkmark
    
End Sub
Sub LoadModel(ProdName As String)
   Dim CurrBkMrk As String
   Dim fldName As String
        lstModel.Clear
    currbkmark = Data2.Recordset.Bookmark
    Data2.Recordset.MoveFirst
    
    Do Until (Data2.Recordset.EOF)
        fldName = Data2.Recordset.Fields(0).Value
        If ProdName = fldName Then
            lstModel.AddItem Data2.Recordset.Fields(0).Value + "  " + Data2.Recordset.Fields(1).Value
        End If
        Data2.Recordset.MoveNext
    Loop
    
    Data2.Recordset.Bookmark = currbkmark
    
End Sub
