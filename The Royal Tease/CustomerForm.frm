VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CustomerForm 
   BackColor       =   &H00400000&
   Caption         =   "Customer Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CustomerForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   14160
      Picture         =   "CustomerForm.frx":82379
      ScaleHeight     =   2595
      ScaleWidth      =   3675
      TabIndex        =   19
      Top             =   5640
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   13560
      Picture         =   "CustomerForm.frx":8469B
      ScaleHeight     =   3915
      ScaleWidth      =   4515
      TabIndex        =   18
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton ButList 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "CustomerForm.frx":8A857
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton ButClose 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Picture         =   "CustomerForm.frx":8B795
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton ButDelete 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "CustomerForm.frx":8C896
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton ButModify 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "CustomerForm.frx":8DA73
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton ButSave 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Picture         =   "CustomerForm.frx":8EC05
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton ButNew 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "CustomerForm.frx":8FD56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5415
      Left            =   9000
      TabIndex        =   16
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   25
      Cols            =   4
      FixedCols       =   0
      GridColorFixed  =   128
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile  No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "CustomerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustVar As String

Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please enter the Customer Name"
Exit Sub
End If

If vbNo = MsgBox("Are you sure you want to Delete this record", vbYesNo, "Delete Record") Then Exit Sub
Conn.Execute "delete from customerTab where custName='" & CustVar & "'"
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
ButList_Click
End Sub

Private Sub ButList_Click()
MSF.Clear
MSF.ColWidth(0) = 2400
MSF.ColWidth(1) = 2500
MSF.TextMatrix(0, 0) = "Customer Name"
MSF.TextMatrix(0, 1) = "Address"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select custName,add3 from CustomerTab order by custname", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0) & ""
MSF.TextMatrix(I, 1) = tRS(1) & ""
I = I + 1
MSF.Rows = I + 5
tRS.MoveNext
Loop


End Sub

Private Sub butModify_Click()
If Text1 = "" Then
MsgBox "Please enter the Customer Name"
Exit Sub
End If
Conn.Execute "delete from customerTab where custName='" & CustVar & "'"
Conn.Execute "insert into CustomerTab values('" & UCase(Text1) & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')"
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
ButList_Click
End Sub

Private Sub ButNew_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""


ButNew.Enabled = False
butSave.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False


End Sub

Private Sub butSave_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
MsgBox "Please enter the Customer Name"
Exit Sub
End If

Conn.Execute "insert into CustomerTab values('" & UCase(Text1) & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')"
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
ButList_Click
End Sub

Private Sub Form_Load()
ButList_Click
End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
CustVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from customerTab where custName='" & CustVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0) & ""
Text2 = tRS(1) & ""
Text3 = tRS(2) & ""
Text4 = tRS(3) & ""
Text5 = tRS(4) & ""
Text6 = tRS(5) & ""
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = True
butDelete.Enabled = True
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub
