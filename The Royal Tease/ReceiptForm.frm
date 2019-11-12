VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ReceiptForm 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Phone Booking Receipt"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "ReceiptForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   12480
      Picture         =   "ReceiptForm.frx":4DFB5
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   17
      Top             =   5160
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H000080FF&
      Height          =   3015
      Left            =   11040
      Picture         =   "ReceiptForm.frx":50880
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   16
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Butclose 
      Height          =   735
      Left            =   9480
      Picture         =   "ReceiptForm.frx":52DA8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton ButDis 
      Height          =   735
      Left            =   7320
      Picture         =   "ReceiptForm.frx":58F24
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton butDelete 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5160
      Picture         =   "ReceiptForm.frx":5F641
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton butModify 
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   735
      Left            =   3120
      Picture         =   "ReceiptForm.frx":65905
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton butSave 
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   735
      Left            =   1560
      Picture         =   "ReceiptForm.frx":6BB4C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton ButNew 
      Height          =   735
      Left            =   120
      Picture         =   "ReceiptForm.frx":71D97
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5895
      Left            =   7080
      TabIndex        =   6
      Top             =   1800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10398
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The Royal Tease Restaurant"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   5880
      TabIndex        =   15
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
   End
End
Attribute VB_Name = "ReceiptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BillingNoVar As String
Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If vbNo = MsgBox("Do you want delete this record", vbYesNo, "Delete Record") Then Exit Sub
Conn.Execute "delete from ReceiptTab where BillingNo=" & BillingNoVar & ""
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButDis_Click()
MSF.Clear
MSF.Cols = 4
MSF.TextMatrix(0, 0) = "BillingNo"
MSF.TextMatrix(0, 1) = "Name"
MSF.TextMatrix(0, 2) = "Amount"

MSF.ColWidth(0) = 1000
MSF.ColWidth(1) = 1200
MSF.ColWidth(2) = 1400

I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from ReceiptTab order by BillingNo", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1)
MSF.TextMatrix(I, 2) = tRS(2)
tRS.MoveNext
I = I + 1
MSF.Rows = I + 5
Loop
End Sub

Private Sub butModify_Click()
If Text2 = "" Then
MsgBox "Please enter the App No"
Exit Sub
End If

Conn.Execute "delete from ReceiptTab where BillingNo=" & BillingNoVar & ""
Conn.Execute "insert into ReceiptTab values('" & Text2 & "','" & Text3 & "'," & Val(Text4) & ",'" & Text5 & "')"


butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButNew_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

butSave.Enabled = True
ButNew.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False

End Sub

Private Sub butSave_Click()
If Text2 = "" Then
MsgBox "Please enter the Billing No"
Exit Sub
End If

Conn.Execute "insert into ReceiptTab values('" & Text2 & "','" & Text3 & "'," & Val(Text4) & ",'" & Text5 & "')"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub



Private Sub Form_Load()

D1 = Date

ButDis_Click

End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
BillingNoVar = MSF.TextMatrix(MSF.Row, 0)
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from ReceiptTab where BillingNo=" & BillingNoVar & "", Conn
If tRS.EOF = False Then
Text2 = tRS(0)
Text3 = tRS(1)
Text4 = tRS(2)
Text5 = tRS(3)


butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = True
butDelete.Enabled = True
End If





End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub

Private Sub Text2_LostFocus()
If tRS.State = 1 Then tRS.Close
tRS.Open "select cName,totAmt from TranmainTab where TranNo=" & Val(Text2) & " and TranType='D'", Conn
If tRS.EOF = False Then

Text3 = tRS(0)
Text4 = tRS(1)
Else
Text3 = ""
Text4 = ""


End If


End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub
