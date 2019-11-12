VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CategoryForm 
   BackColor       =   &H00004000&
   Caption         =   "Category Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CategoryForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   5895
      Left            =   6240
      Picture         =   "CategoryForm.frx":56BE8
      ScaleHeight     =   5835
      ScaleWidth      =   9555
      TabIndex        =   10
      Top             =   1440
      Width           =   9615
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   14880
      ScaleHeight     =   15
      ScaleWidth      =   1335
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
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
      Left            =   15840
      Picture         =   "CategoryForm.frx":72F1B
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   18240
      Picture         =   "CategoryForm.frx":78AFB
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "CategoryForm.frx":7EF73
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   2055
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
      Left            =   3600
      Picture         =   "CategoryForm.frx":8516C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   1935
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
      Left            =   1800
      Picture         =   "CategoryForm.frx":8B28A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1695
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
      Left            =   0
      Picture         =   "CategoryForm.frx":9113F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1695
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
      Top             =   3600
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
      Top             =   2640
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6975
      Left            =   16440
      TabIndex        =   8
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   12303
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
      Caption         =   "The Royal Tease restaurant"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   11055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "CategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CatVar As String

Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please enter the Category Name"
Exit Sub
End If

If vbNo = MsgBox("Are you sure you want to Delete this record", vbYesNo, "Delete Record") Then Exit Sub
Conn.Execute "delete from CategoryTab where CategoryName='" & CatVar & "'"
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
MSF.TextMatrix(0, 0) = "CategoryName"
MSF.TextMatrix(0, 1) = "Details"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select CategoryName,Details from CategoryTab order by CategoryName", Conn
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
MsgBox "Please enter the Category Name"
Exit Sub
End If
Conn.Execute "delete from CategoryTab where CategoryName='" & CatVar & "'"
Conn.Execute "insert into CategoryTab values('" & UCase(Text1) & "','" & Text2 & "')"
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
ButList_Click
End Sub

Private Sub ButNew_Click()
Text1 = ""
Text2 = ""



ButNew.Enabled = False
butSave.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False


End Sub

Private Sub butSave_Click()
If Text1 = "" Or Text2 = "" Then
MsgBox "Please fill details"
Exit Sub
End If

Conn.Execute "insert into CategoryTab values('" & UCase(Text1) & "','" & Text2 & "')"
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
CatVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CategoryTab where CategoryName='" & CatVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0) & ""
Text2 = tRS(1) & ""

ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = True
butDelete.Enabled = True
End If
End Sub
