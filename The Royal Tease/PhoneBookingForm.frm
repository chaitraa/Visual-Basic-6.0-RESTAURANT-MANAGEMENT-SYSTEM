VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PhoneBookingForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   Caption         =   "Phone Booking Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PhoneBookingForm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      Height          =   390
      Left            =   480
      TabIndex        =   30
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      MaxLength       =   10
      TabIndex        =   28
      Top             =   5400
      Width           =   3975
   End
   Begin VB.CommandButton butPrint 
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
      Left            =   15600
      Picture         =   "PhoneBookingForm.frx":75A94
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox TextAdd 
      Height          =   2055
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   2760
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      Height          =   390
      Left            =   480
      TabIndex        =   24
      Top             =   2160
      Width           =   3975
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
      Left            =   480
      Picture         =   "PhoneBookingForm.frx":7B86D
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8400
      Width           =   1575
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
      Left            =   2040
      Picture         =   "PhoneBookingForm.frx":815B3
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8400
      Width           =   1575
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
      Left            =   3600
      Picture         =   "PhoneBookingForm.frx":87468
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8400
      Width           =   1815
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
      Picture         =   "PhoneBookingForm.frx":8D586
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8400
      Width           =   1935
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
      Left            =   17880
      Picture         =   "PhoneBookingForm.frx":9377F
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8160
      Width           =   1815
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
      Left            =   7440
      Picture         =   "PhoneBookingForm.frx":99BF7
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox TextgTotal 
      Height          =   375
      Left            =   17640
      MaxLength       =   50
      TabIndex        =   15
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton butAdd 
      Caption         =   "Add"
      Height          =   615
      Left            =   19080
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextTotal 
      Height          =   375
      Left            =   17760
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox TextQty 
      Height          =   375
      Left            =   16800
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   12720
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox TextRate 
      Height          =   375
      Left            =   15480
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   480
      Left            =   6840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   480
      Left            =   480
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker bDate 
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   95617027
      CurrentDate     =   40481
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4335
      Left            =   12720
      TabIndex        =   6
      Top             =   2880
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   25
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "The Royal tease Restaurant"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8520
      TabIndex        =   32
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned to Emp"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   12960
      TabIndex        =   23
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   16320
      TabIndex        =   16
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   17760
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   16920
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   15600
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Bill No"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Ref No"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "PhoneBookingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rowNoVar, tranVar As Long
Dim tAmtVar As Long

Private Sub butAdd_Click()
If Val(TextQty) <= 0 Then
MsgBox "Please enter the Quantity"
Exit Sub
End If

MSF.TextMatrix(rowNoVar, 0) = rowNoVar
MSF.TextMatrix(rowNoVar, 1) = Combo1
MSF.TextMatrix(rowNoVar, 2) = TextRate
MSF.TextMatrix(rowNoVar, 3) = TextQty
MSF.TextMatrix(rowNoVar, 4) = TextTotal
rowNoVar = rowNoVar + 1
tAmtVar = tAmtVar + Val(TextTotal)
Combo1 = ""
TextQty = ""
TextTotal = ""
TextgTotal = ""
TextgTotal = tAmtVar
End Sub

Private Sub Butclose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If Text1 = "" Then
MsgBox "RefNo is not valid"
Exit Sub
End If

If vbNo = MsgBox("Are you sure you want to Delete this record", vbYesNo, "Delete Record") Then Exit Sub
Conn.Execute "delete from TranMainTab where TranNo=" & tranVar & ""
Conn.Execute "delete from TranDetailTab where TranNo=" & tranVar & ""
ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub ButList_Click()
tranVar = InputBox("Please Enter the Sales Ref No")
If Val(tranVar) <= 0 Then Exit Sub

MSFInit

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from TranMainTab where TranNo=" & tranVar & "", Conn
If tRS.EOF = False Then
Text1 = tRS(0)
bDate = tRS(1)
Text2 = tRS(3)
Combo2 = tRS(4)
tAmtVar = tRS(8)
TextgTotal = tAmtVar
End If


If tRS.State = 1 Then tRS.Close
tRS.Open "select * from TranDetailTab where TranNo=" & tranVar & "", Conn
Do While tRS.EOF = False
MSF.TextMatrix(rowNoVar, 0) = rowNoVar
MSF.TextMatrix(rowNoVar, 1) = tRS(2)
MSF.TextMatrix(rowNoVar, 2) = tRS(3)
MSF.TextMatrix(rowNoVar, 3) = tRS(4)
MSF.TextMatrix(rowNoVar, 4) = tRS(5)
rowNoVar = rowNoVar + 1
MSF.Rows = rowNoVar + 5
tRS.MoveNext
Loop

ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = True
butDelete.Enabled = True

End Sub

Private Sub butModify_Click()
If vbNo = MsgBox("Are you sure you want to Modify this record", vbYesNo, "Modify Record") Then Exit Sub
Conn.Execute "delete from TranMainTab where TranNo=" & tranVar & ""
Conn.Execute "delete from TranDetailTab where TranNo=" & tranVar & ""

butSave_Click

End Sub

Private Sub ButNew_Click()
Text1 = ""
Text2 = ""
TextRate = ""
TextQty = ""
TextTotal = ""
TextgTotal = ""
TextAdd = ""

MSFInit

If tRS.State = 1 Then tRS.Close
tRS.Open "select max(tranNo) from TranMainTab", Conn
Text1 = IIf(IsNull(tRS(0)), 1000, tRS(0)) + 1

If tRS.State = 1 Then tRS.Close
tRS.Open "select max(BillNo) from TranMainTab where tranType='S'", Conn
Text2 = IIf(IsNull(tRS(0)), 1000, tRS(0)) + 1

ButNew.Enabled = False
butSave.Enabled = True
butModify.Enabled = False
butDelete.Enabled = False

End Sub

Private Sub butPrint_Click()
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from TranDetailTab where TranNo=" & Val(Text1) & "", Conn
Set SalesbillPrint.DataSource = tRS
SalesbillPrint.Sections("section2").Controls("L1").Caption = Combo2
SalesbillPrint.Sections("section2").Controls("L2").Caption = TextAdd
SalesbillPrint.Sections("section2").Controls("L3").Caption = Text2
SalesbillPrint.Sections("section2").Controls("L4").Caption = DateFormat(bDate)
SalesbillPrint.Sections("section3").Controls("L5").Caption = TextgTotal
SalesbillPrint.Show
End Sub

Private Sub butSave_Click()
If Combo2 = "" Then
MsgBox "Please enter the Supplier Name"
Exit Sub
End If

Conn.Execute "insert into TranMainTab values(" & Val(Text1) & ",'" & DateFormat(bDate) & "','D'," & Val(Text2) & ",'" & Combo2 & "','" & Text3 & "','" & Text4 & "','" & Combo3 & "'," & Val(TextgTotal) & ")"
For I = 1 To rowNoVar - 1
Conn.Execute "insert into TranDetailTab values(" & Val(Text1) & "," & Val(MSF.TextMatrix(I, 0)) & ",'" & MSF.TextMatrix(I, 1) & "'," & Val(MSF.TextMatrix(I, 2)) & "," & Val(MSF.TextMatrix(I, 3)) & "," & Val(MSF.TextMatrix(I, 4)) & ")"
Next

ButNew.Enabled = True
butSave.Enabled = False
butModify.Enabled = False
butDelete.Enabled = False
End Sub

Private Sub Combo1_LostFocus()
If tRS.State = 1 Then tRS.Close
tRS.Open "select sPrice from ProductTab where ProdCode='" & Combo1 & "'", Conn
If tRS.EOF = False Then
TextRate = tRS(0)
Else
TextRate = ""
End If
End Sub

Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 600
MSF.ColWidth(1) = 3600
MSF.ColWidth(2) = 1200
MSF.ColWidth(3) = 1200
MSF.ColWidth(4) = 1600
MSF.TextMatrix(0, 0) = "SlNo"
MSF.TextMatrix(0, 1) = "Product Code"
MSF.TextMatrix(0, 2) = "Rate"
MSF.TextMatrix(0, 3) = "Qty"
MSF.TextMatrix(0, 4) = "Total"
rowNoVar = 1
tAmtVar = 0
End Sub
Private Sub Combo2_LostFocus()
If tRS.State = 1 Then tRS.Close
tRS.Open "select Add1,add2,add3,pincode from customerTab where custName='" & Combo2 & "'", Conn
If tRS.EOF = False Then
TextAdd = tRS(0) & "" & vbCrLf
TextAdd = TextAdd & tRS(1) & "" & vbCrLf
TextAdd = TextAdd & tRS(2) & "" & vbCrLf
TextAdd = TextAdd & tRS(3) & "" & vbCrLf
Else
TextAdd = ""
End If
End Sub

Private Sub Form_Load()
bDate = Date
MSFInit
If tRS.State = 1 Then tRS.Close
tRS.Open "select custName from customerTab order by custName", Conn
Do While tRS.EOF = False
Combo2.AddItem tRS(0)
tRS.MoveNext
Loop

If tRS.State = 1 Then tRS.Close
tRS.Open "select Prodcode from ProductTab order by prodcode", Conn
Do While tRS.EOF = False
Combo1.AddItem tRS(0)
tRS.MoveNext
Loop

If tRS.State = 1 Then tRS.Close
tRS.Open "select empCode from empTab order by empcode", Conn
Do While tRS.EOF = False
Combo3.AddItem tRS(0)
tRS.MoveNext
Loop
End Sub

Private Sub MSF_DblClick()
Dim ftNo, ftNo1 As Long
If Not MSF.TextMatrix(MSF.Row, 1) = "" Then
If vbNo = MsgBox("Do you want to remove this Item", vbYesNo, "Remove") Then Exit Sub
ftNo = MSF.Row
ftNo1 = (rowNoVar - 1)
tAmtVar = tAmtVar - MSF.TextMatrix(MSF.Row, 4)
        For I = ftNo To ftNo1
        MSF.TextMatrix(I, 0) = I
        MSF.TextMatrix(I, 1) = MSF.TextMatrix(I + 1, 1)
        MSF.TextMatrix(I, 2) = MSF.TextMatrix(I + 1, 2)
        MSF.TextMatrix(I, 3) = MSF.TextMatrix(I + 1, 3)
        MSF.TextMatrix(I, 4) = MSF.TextMatrix(I + 1, 4)
        Next I
        MSF.TextMatrix(ftNo1, 0) = ""
rowNoVar = rowNoVar - 1
TextgTotal = tAmtVar
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub

Private Sub TextQty_Change()
TextTotal = Val(TextRate) * Val(TextQty)
End Sub

Private Sub TextQty_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub

Private Sub TextRate_Change()
TextTotal = Val(TextRate) * Val(TextQty)
End Sub

Private Sub TextRate_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNum(KeyAscii)
End Sub
