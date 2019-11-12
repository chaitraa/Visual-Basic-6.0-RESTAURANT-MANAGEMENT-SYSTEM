VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EmpListForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Emp List Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "EmpListForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
      Height          =   735
      Left            =   10680
      Picture         =   "EmpListForm.frx":4B23F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      Picture         =   "EmpListForm.frx":4D05F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      Picture         =   "EmpListForm.frx":4EB5E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8175
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   25
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp List"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Royal Tease Restaurant"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "EmpListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSF.Clear
MSF.Cols = 6
MSF.ColWidth(0) = 1400
MSF.ColWidth(1) = 1600
MSF.ColWidth(2) = 2000
MSF.ColWidth(3) = 3400
MSF.ColWidth(4) = 1400
MSF.TextMatrix(0, 0) = "Emp Code"
MSF.TextMatrix(0, 1) = "Name"
MSF.TextMatrix(0, 2) = "Address"
MSF.TextMatrix(0, 3) = "Qualification"
MSF.TextMatrix(0, 4) = "Designation"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from EmpTab", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1) & ""
MSF.TextMatrix(I, 2) = tRS(2) & ""
MSF.TextMatrix(I, 3) = tRS(3) & ""
MSF.TextMatrix(I, 4) = tRS(4) & ""
I = I + 1
MSF.Rows = I + 5
tRS.MoveNext
Loop

End Sub

Private Sub Command2_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from EmpTab order by EmpCode", Conn
Set EmpListReport.DataSource = tRS
EmpListReport.Show
End Sub


Private Sub Command5_Click()
Unload Me
End Sub

