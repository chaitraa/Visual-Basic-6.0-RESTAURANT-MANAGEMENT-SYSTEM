VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CustomerListReport 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Customer List Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CustomerListReport.frx":0000
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
      Height          =   615
      Left            =   10800
      Picture         =   "CustomerListReport.frx":82379
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
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
      Height          =   615
      Left            =   10800
      Picture         =   "CustomerListReport.frx":84199
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
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
      Height          =   735
      Left            =   11040
      Picture         =   "CustomerListReport.frx":85C98
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8175
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   25
      Cols            =   3
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " The Royal Tease Restaurant"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "CustomerListReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSF.Clear
MSF.ColWidth(0) = 2400
MSF.ColWidth(1) = 6000
MSF.ColWidth(2) = 1800
MSF.TextMatrix(0, 0) = "Name"
MSF.TextMatrix(0, 1) = "Address"
MSF.TextMatrix(0, 2) = "Phone"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CustomerTab order by custname", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1) & ", " & tRS(2) & ", " & tRS(3) & ", " & tRS(4) & ", "
MSF.TextMatrix(I, 2) = tRS(5) & ""
I = I + 1
MSF.Rows = I + 5
tRS.MoveNext
Loop

End Sub

Private Sub Command2_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from CustomerTab order by custname", Conn
Set CustomerListPrint.DataSource = tRS
CustomerListPrint.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
