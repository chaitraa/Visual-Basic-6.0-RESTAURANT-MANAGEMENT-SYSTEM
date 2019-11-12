VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BookingReportForm 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Booking Sales List Report"
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
   Picture         =   "BookingReportForm.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker fDate 
      Height          =   495
      Left            =   10800
      TabIndex        =   4
      Top             =   600
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
      Format          =   95748099
      CurrentDate     =   40481
   End
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   11040
      Picture         =   "BookingReportForm.frx":50D99
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3855
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
      Left            =   11040
      Picture         =   "BookingReportForm.frx":51FA4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   3855
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
      Picture         =   "BookingReportForm.frx":52DBC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   13996
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
      Left            =   10800
      TabIndex        =   7
      Top             =   1560
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
      Format          =   95748099
      CurrentDate     =   40481
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "BookingReportForm"
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
tRS.Open "select * from TranMainTab where tranType='D' and (trandate >=#" & DateFormat(fDate) & "# and trandate <=#" & DateFormat(tDate) & "#) order by TranNo", Conn
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
tRS.Open "select * from TranMainTab where tranType='D' and (trandate >=#" & DateFormat(fDate) & "# and trandate <=#" & DateFormat(tDate) & "#) order by TranNo", Conn
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
