VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EmpForm 
   BackColor       =   &H00404000&
   Caption         =   "Employee  Details Form"
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
   Picture         =   "EmpForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   13320
      Picture         =   "EmpForm.frx":48293
      ScaleHeight     =   2475
      ScaleWidth      =   4275
      TabIndex        =   20
      Top             =   8040
      Width           =   4335
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   15600
      Picture         =   "EmpForm.frx":4BABD
      ScaleHeight     =   2595
      ScaleWidth      =   3795
      TabIndex        =   18
      Top             =   5280
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   12840
      Picture         =   "EmpForm.frx":4DFE3
      ScaleHeight     =   3315
      ScaleWidth      =   4275
      TabIndex        =   17
      Top             =   1800
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "EmpForm.frx":52168
      Left            =   2160
      List            =   "EmpForm.frx":52178
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      ItemData        =   "EmpForm.frx":52198
      Left            =   2160
      List            =   "EmpForm.frx":521A8
      TabIndex        =   5
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   1200
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton Butclose 
      Height          =   615
      Left            =   9000
      Picture         =   "EmpForm.frx":521D2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton ButDis 
      Height          =   615
      Left            =   7080
      Picture         =   "EmpForm.frx":5864A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton butDelete 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      Picture         =   "EmpForm.frx":5E47E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton butModify 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "EmpForm.frx":64677
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton butSave 
      BackColor       =   &H00004080&
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      MaskColor       =   &H0000C000&
      Picture         =   "EmpForm.frx":6A6C7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton ButNew 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   0
      Picture         =   "EmpForm.frx":7057C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5775
      Left            =   7680
      TabIndex        =   11
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   10186
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " The Royal Tease Restaurant"
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
      Height          =   735
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Code"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "EmpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eCodeVar As String
Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "Please Enter The Employee Code"
Exit Sub
End If

Conn.Execute "delete from EmpTab where EmpCode ='" & eCodeVar & "'"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButDis_Click()
MSF.Clear
MSF.Cols = 2
MSF.TextMatrix(0, 0) = "Emp Code"
MSF.TextMatrix(0, 1) = "Emp Name"

MSF.ColWidth(0) = 2000
MSF.ColWidth(1) = 2000

I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from EmpTab  order by EmpCode", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1)
tRS.MoveNext
I = I + 1
MSF.Rows = I + 5
Loop
End Sub

Private Sub butModify_Click()
If Text1 = "" Then
MsgBox "Please Enter The Employee Code"
Exit Sub
End If

Conn.Execute "delete from EmpTab where EmpCode ='" & eCodeVar & "'"
Conn.Execute "insert into EmpTab values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "','" & Left(Combo2, 4) & "')"

butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButNew_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

butSave.Enabled = True
ButNew.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False

End Sub

Private Sub butSave_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Combo2 = "" Then
MsgBox "Please fill all the details"
Exit Sub
End If


Conn.Execute "insert into EmpTab values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "','" & Left(Combo2, 4) & "')"
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False
End Sub






Private Sub Form_Load()


ButDis_Click

End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
eCodeVar = MSF.TextMatrix(MSF.Row, 0)

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from EmpTab where EmpCode ='" & eCodeVar & "'", Conn
If tRS.EOF = False Then
Text1 = tRS(0)
Text2 = tRS(1)
Text3 = tRS(2)
Combo1 = tRS(3)
Combo2 = tRS(4)
End If
butSave.Enabled = False
ButNew.Enabled = True
butModify.Enabled = True
butDelete.Enabled = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
ElseIf (KeyAscii < 65 And KeyAscii <> 8 And KeyAscii <> 32) Or (KeyAscii > 90 And KeyAscii < 97) Or (KeyAscii > 122) Then
KeyAscii = 0
MsgBox ("Please Enter char")
End If

End Sub
