VERSION 5.00
Begin VB.Form LoginForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LoginForm.frx":0000
   ScaleHeight     =   10245
   ScaleWidth      =   17145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Picture         =   "LoginForm.frx":177D7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Picture         =   "LoginForm.frx":186CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from LoginTab where LoginName='" & Text1 & "' and Password1 ='" & Text2 & "'", Conn
If tRS.EOF = False Then
Unload Me
MDIForm1.Show

Else
MsgBox "Entered LoginUserName or Password is not correct Please check"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Exit Sub
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

End Sub

Private Sub Form_Load()
Me.Height = 5175
Me.Width = 7440
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RestuarantData.mdb;Persist Security Info=False"

End Sub

Private Sub Text1_Change()

End Sub
