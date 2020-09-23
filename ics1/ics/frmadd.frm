VERSION 5.00
Begin VB.Form frmadd 
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmp 
      Caption         =   "Move  Previous"
      Height          =   375
      Left            =   4320
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   24
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtmc 
      Alignment       =   1  'Right Justify
      DataField       =   "model"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   " "
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Main"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add New"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdml 
      Caption         =   "Move Last"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdmn 
      Caption         =   "Move Next"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdmf 
      Caption         =   "Move First"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      DataField       =   "quantity"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   " "
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtpd 
      Alignment       =   1  'Right Justify
      DataField       =   "date"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   " "
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtrp 
      Alignment       =   1  'Right Justify
      DataField       =   "reprice"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtcp 
      Alignment       =   1  'Right Justify
      DataField       =   "cost"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Tag             =   "NUMBERIC"
      Text            =   " "
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtpn 
      Alignment       =   1  'Right Justify
      DataField       =   "product"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\VBPROJECT\ics\db1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ics"
      Top             =   840
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   3720
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   4200
      Y1              =   840
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   5760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5760
      Y1              =   840
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   5760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label10 
      Caption         =   " "
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   " "
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Add Record"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
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
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Production Date"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Retail Price"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cost Price"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Model Code"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rec As Recordset
Public eno As Boolean

Private Sub cmdcancel_Click()
Data4.Recordset.CancelUpdate
cmddelete.Enabled = True
cmdexit.Enabled = True
cmdedit.Enabled = True
cmdmf.Enabled = True
cmdml.Enabled = True
cmdmn.Enabled = True
cmdadd.Enabled = True
cmdmp.Enabled = True
cmdcancel.Enabled = False
txtpd.Locked = True
txtpn.Locked = True
txtcp.Locked = True
txtmc.Locked = True
txtrp.Locked = True
txtqty.Locked = True
cmdupdate.Enabled = False

End Sub

Private Sub cmddelete_Click()
Dim response As Integer
       response = MsgBox("Are you sure you want to delete?", vbYesNo + vbQuestion, "Are you Sure?")
        If response = vbYes Then
        With Data4.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
        MsgBox "Record Deleted!!", vbInformation
    Else
        MsgBox "Record Not Delete!!", vbInformation
    Exit Sub
        End If
   End Sub

Private Sub cmdexit_Click()
Unload frmadd
frmmain.Show
End Sub

Private Sub cmdadd_Click()
Data4.Recordset.AddNew
cmdcancel.Enabled = True
cmdadd.Enabled = False
cmddelete.Enabled = False
cmdexit.Enabled = False
cmdedit.Enabled = False
cmdmf.Enabled = False
cmdml.Enabled = False
cmdmn.Enabled = False
cmdmp.Enabled = False
cmdupdate.Enabled = True
'txtpd.Locked = False
'txtpn.Locked = False
'txtcp.Locked = False
'txtmc.Locked = False
'txtrp.Locked = False
'txtqty.Locked = False
txtpd.Text = Date
txtpn.SetFocus
End Sub

Private Sub cmdedit_Click()
msg = MsgBox("YOU'RE EDIT AN RECORD !", vbOKCancel, "Warnning")
If msg = vbOK Then
Data4.Recordset.Edit
cmddelete.Enabled = False
cmdadd.Enabled = False
cmdexit.Enabled = False
cmdmf.Enabled = False
cmdml.Enabled = False
cmdmn.Enabled = False
cmdmp.Enabled = False
cmdedit.Enabled = False
cmdcancel.Enabled = True
cmdupdate.Enabled = True
txtpd.Locked = True
txtpn.Locked = True
txtcp.Locked = True
txtmc.Locked = True
txtrp.Locked = True
txtqty.Locked = True
Else
frmadd.Data4.Recordset.Edit
frmadd.Data4.Recordset.CancelUpdate
MsgBox "Action cancel"
End If
End Sub

Private Sub cmdmf_Click()
Data4.Recordset.MoveFirst
End Sub

Private Sub cmdml_Click()
Data4.Recordset.MoveLast
End Sub

Private Sub cmdmn_Click()
Data4.Recordset.MoveNext
If Data4.Recordset.EOF Then
Data4.Recordset.MoveLast
End If
End Sub

Private Sub cmdmp_Click()
Data4.Recordset.MovePrevious
If Data4.Recordset.BOF Then
Data4.Recordset.MoveFirst
End If
End Sub

Private Sub cmdupdate_Click()
If (txtpn.Text) <> "" Then
msg = MsgBox("ARE YOU SURE WANT TO UPDATE ?", vbOKCancel + vbQuestion, "Update")
If msg = vbOK Then
Data4.Recordset.Update
cmdadd.Enabled = True
cmddelete.Enabled = True
cmdexit.Enabled = True
cmdedit.Enabled = True
cmdmf.Enabled = True
cmdml.Enabled = True
cmdmp.Enabled = True
cmdmn.Enabled = True
cmdcancel.Enabled = False
cmdupdate.Enabled = False
txtpd.Locked = True
txtpn.Locked = True
txtcp.Locked = True
txtmc.Locked = True
txtrp.Locked = True
txtqty.Locked = True
Else
cmdcancel.SetFocus
End If
Else
MsgBox "Record is blank!!"
cmdcancel.SetFocus
End If

End Sub

Private Sub Form_Load()
'Set db = OpenDatabase(App.Path & "\db1.mdb")
'Set rec = db.OpenRecordset("select * from ics")

cmdupdate.Enabled = False
Label9.Caption = Date
Label10.Caption = Time
End Sub


Private Sub txtcp_Click()
If txtpn.Text = Empty Then
MsgBox "PRODUCT NAME IS EMPTY"
txtpn.SetFocus
End If
End Sub




Private Sub txtcp_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = (KeyAscii > 45 And KeyAscii < 58) Or (KeyAscii = 8) Or (KeyAscii = 13)
'txtcp.Locked = True
If eno = True Then
txtcp.Locked = False
If KeyAscii = 13 Then
If Not (txtcp.Text) = "" Then
txtrp.SetFocus
'Text1.Text = UCase(Text1.Text)
Else
txtcp.Locked = True
MsgBox "THE COST PRICE IS EMPTY"
txtcp.SetFocus
txtcp.Text = ""
End If
End If
End If
End Sub

Private Sub txtmc_Click()
If txtpn.Text = "" Then
MsgBox " YOU MUST ENTER THE PRODUCT NAME FRIST"
txtpn.SetFocus
Else
txtmc.SetFocus
End If
End Sub

Private Sub txtmc_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = Chr(KeyAscii) Like "[A-Za-z]" Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 13
'txtmc.Locked = True
If eno = True Then
txtmc.Locked = False
If KeyAscii = 13 Then
If Not (txtmc.Text) = "" Then
txtcp.SetFocus
txtmc.Text = UCase(txtmc.Text)
Else
MsgBox "THE MODEL CODE IS EMPTY"
txtmc.SetFocus
txtmc.Text = ""
End If
End If
End If
End Sub

Private Sub txtpd_Click()
If txtpn.Text <> Empty Then
txtpd.SetFocus
Else
MsgBox "Product Name is empty!!"
txtpn.SetFocus
End If
End Sub

Private Sub txtpd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsDate(txtpd.Text) = True Then ' Isdate is a command to use to check the date whether is in the correct format
txtqty.SetFocus
Else
MsgBox "INVALID DATE ( EXAMPLE : 01/12/01 )", vbCritical
txtpd.Text = Date
txtpd.SetFocus
End If
End If
End Sub



Private Sub txtpn_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 13
'txtpn.Locked = True
If eno = True Then
txtpn.Locked = False
If KeyAscii = 13 Then
If Not (txtpn.Text) = "" Then
txtmc.Enabled = True
txtmc.SetFocus
txtpn.Text = UCase(txtpn.Text)
Else
KeyAscii = 0
MsgBox "THE PRODUCT NAME IS EMPTY"
txtpn.SetFocus
txtpn.Text = ""
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub

Private Sub txtqty_Click()
If txtpn.Text = Empty Then
MsgBox "PRODUCT NAME IS EMPTY"
txtpn.SetFocus
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Or (KeyAscii = 13)
'txtqty.Locked = True
If eno = True Then
txtqty.Locked = False
If KeyAscii = 13 Then
If Not (txtqty.Text) = "" Then
cmdupdate.SetFocus
'Text1.Text = UCase(Text1.Text)
Else
txtqty.Locked = True
MsgBox "THE QUANTITY IS EMPTY"
txtqty.SetFocus
txtqty.Text = ""
End If
End If
End If
End Sub

Private Sub txtrp_Click()
If txtpn.Text = Empty Then
MsgBox "PRODUCT NAME IS EMPTY"
txtpn.SetFocus
End If
End Sub

Private Sub txtrp_KeyPress(KeyAscii As Integer)
Dim eno As Boolean
eno = (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Or (KeyAscii = 13)
'txtrp.Locked = True
If eno = True Then
txtrp.Locked = False
If KeyAscii = 13 Then
If Not (txtrp.Text) = "" Then
txtpd.SetFocus
'Text1.Text = UCase(Text1.Text)
Else
txtrp.Locked = True
MsgBox "THE REPRICE IS EMPTY"
txtrp.SetFocus
txtrp.Text = ""
End If
End If
End If
End Sub
