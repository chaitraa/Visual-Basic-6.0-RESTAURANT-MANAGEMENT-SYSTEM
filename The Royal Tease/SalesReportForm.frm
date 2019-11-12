VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesReportForm 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Sales List Report"
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
   Picture         =   "SalesReportForm.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   7680
      Picture         =   "SalesReportForm.frx":47B26
      ScaleHeight     =   2835
      ScaleWidth      =   3075
      TabIndex        =   10
      Top             =   3000
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker fDate 
      Height          =   495
      Left            =   11640
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   95617027
      CurrentDate     =   40481
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
      Left            =   12240
      Picture         =   "SalesReportForm.frx":4A3F1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
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
      Height          =   495
      Left            =   12240
      Picture         =   "SalesReportForm.frx":4B0F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
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
      Height          =   495
      Left            =   12360
      Picture         =   "SalesReportForm.frx":4BBE3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8175
      Left            =   480
      TabIndex        =   3
      Top             =   1200
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
   Begin MSComCtl2.DTPicker tDate 
      Height          =   495
      Left            =   11640
      TabIndex        =   7
      Top             =   3240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   95617027
      CurrentDate     =   40481
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "sales list"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   11640
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   11640
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "SalesReportForm"
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
MSF.TextMatrix(0, 0) = "RefNo"
MSF.TextMatrix(0, 1) = "Date"
MSF.TextMatrix(0, 2) = "BillNo"
MSF.TextMatrix(0, 3) = "Party Name"
MSF.TextMatrix(0, 4) = "Amount"
I = 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from TranMainTab where tranType='S' and (trandate >=#" & DateFormat(fDate) & "# and trandate <=#" & DateFormat(tDate) & "#) order by TranNo", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS(0)
MSF.TextMatrix(I, 1) = tRS(1) & ""
MSF.TextMatrix(I, 2) = DateFormat(tRS(3)) & ""
MSF.TextMatrix(I, 3) = tRS(4) & ""
MSF.TextMatrix(I, 4) = tRS.Fields("totAmt") & ""
I = I + 1
MSF.Rows = I + 5
tRS.MoveNext
Loop

End Sub

Private Sub Command2_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from TranMainTab where tranType='S' and (trandate >=#" & DateFormat(fDate) & "# and trandate <=#" & DateFormat(tDate) & "#) order by TranNo", Conn
Set SalesReport.DataSource = tRS
SalesReport.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
fDate = Date
tDate = Date

End Sub

