VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ProductForm 
   BackColor       =   &H00404000&
   Caption         =   "Product Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "ProductForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture6 
      Height          =   2175
      Left            =   17400
      Picture         =   "ProductForm.frx":48293
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   25
      Top             =   960
      Width           =   2775
   End
   Begin VB.PictureBox Picture5 
      Height          =   2655
      Left            =   15000
      Picture         =   "ProductForm.frx":4A45D
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   24
      Top             =   6360
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   11400
      Picture         =   "ProductForm.frx":4C4AA
      ScaleHeight     =   2355
      ScaleWidth      =   3435
      TabIndex        =   23
      Top             =   3600
      Width           =   3495
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   12720
      Picture         =   "ProductForm.frx":4FC57
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   22
      Top             =   840
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   11760
      Picture         =   "ProductForm.frx":53BCA
      ScaleHeight     =   2475
      ScaleWidth      =   2835
      TabIndex        =   21
      Top             =   6240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   2535
      Left            =   15360
      Picture         =   "ProductForm.frx":564AD
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   20
      Top             =   3600
      Width           =   3975
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
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
      Height          =   735
      Left            =   7560
      Picture         =   "ProductForm.frx":589E4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8160
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
      Height          =   735
      Left            =   9240
      Picture         =   "ProductForm.frx":5E5C4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   1935
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
      Height          =   735
      Left            =   5640
      Picture         =   "ProductForm.frx":64A3C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
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
      Height          =   735
      Left            =   3720
      Picture         =   "ProductForm.frx":6AC35
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
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
      Height          =   735
      Left            =   2040
      Picture         =   "ProductForm.frx":70D53
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
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
      Height          =   735
      Left            =   360
      Picture         =   "ProductForm.frx":76C08
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
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
      Height          =   510
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2250
      Width           =   3255
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6255
      Left            =   7200
      TabIndex        =   16
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   25
      Cols            =   4
      FixedCols       =   0
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   2295
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
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6840
      TabIndex        =   17
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "ProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProdVar As String
Function CheckNum(KeyNum)
If KeyNum = 8 Then CheckNum = KeyNum: Exit Function
If KeyNum < 46 Or KeyNum > 57 Then
CheckNum = 0
MsgBox ("Please Enter Numbers Only")
Else
CheckNum = KeyNum
End If
End Function


Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please enter the Product Code"
Exit Sub
End If

If vbNo = MsgBox("Are you sure you want to Delete this record", vbYesNo, "Delete Record") Then Exit Sub
Conn.Execute "delete from ProductTab where ProdCode='" & ProdVar & "'"
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
MSF.TextMatrix(0, 0) = "Product Code"
MSF.TextMatrix(0, 1) = "Category"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select ProdCode,pCat from ProductTab order by ProdCode", Conn
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
MsgBox "Please enter the Product Code"
Exit Sub
End If
Conn.Execute "delete from ProductTab where ProdCode='" & ProdVar & "'"
Conn.Execute "insert into ProductTab values('" & UCase(Text1) & "','" & Text2 & "','" & Combo1 & "'," & Val(Text3) & "," & Val(Text4) & ",'" & Text5 & "')"
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
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Combo1 = "" Then
MsgBox "Please enter complete details"
Exit Sub
End If

Conn.Execute "insert into ProductTab values('" & UCase(Text1) & "','" & Text2 & "','" & Combo1 & "'," & Val(Text3) & "," & Val(Text4) & ",'" & Text5 & "')"
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
ButList_Click
End Sub

Private Sub Form_Load()
ButList_Click
If tRS.State = 1 Then tRS.Close
tRS.Open "select CategoryName,Details from CategoryTab order by CategoryName", Conn
Do While tRS.EOF = False
Combo1.AddItem (tRS(0))
tRS.MoveNext
Loop
End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
ProdVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from ProductTab where ProdCode='" & ProdVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0) & ""
Text2 = tRS(1) & ""
Combo1 = tRS(2) & ""
Text3 = tRS(3) & ""
Text4 = tRS(4) & ""
Text5 = tRS(5) & ""

ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = True
butDelete.Enabled = True
End If
End Sub




Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub




Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
ElseIf (KeyAscii < 65 And KeyAscii <> 8 And KeyAscii <> 32) Or (KeyAscii > 90 And KeyAscii < 97) Or (KeyAscii > 122) Then
KeyAscii = 0
MsgBox ("Please Enter char")
End If
End Sub
