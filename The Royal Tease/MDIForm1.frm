VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404000&
   Caption         =   "Restuarant Mng"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   1455
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   10095
         TabIndex        =   2
         Top             =   600
         Width           =   10095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu CatMenu 
      Caption         =   "Category"
   End
   Begin VB.Menu ProductMenu 
      Caption         =   "Products"
   End
   Begin VB.Menu EmpMenu 
      Caption         =   "Employee"
   End
   Begin VB.Menu CustMenu 
      Caption         =   "Customer"
   End
   Begin VB.Menu BillingMenu 
      Caption         =   "Billing"
   End
   Begin VB.Menu BillingRepMenu 
      Caption         =   "Billing Report"
   End
   Begin VB.Menu PhonebookMenu 
      Caption         =   "PhoneBooking"
   End
   Begin VB.Menu DelRecMenu 
      Caption         =   "Delivery Receipt"
   End
   Begin VB.Menu ReportsMainMenu 
      Caption         =   "Reports"
      Begin VB.Menu EmpListRepMenu 
         Caption         =   "Employee List"
      End
      Begin VB.Menu CustomerListMenu 
         Caption         =   "Customer List"
      End
      Begin VB.Menu DelRecListMenu 
         Caption         =   "Delivery Receipts  List"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BillingMenu_Click()
SalesForm.Show
End Sub

Private Sub BillingRepMenu_Click()
SalesReportForm.Show
End Sub

Private Sub CatMenu_Click()
CategoryForm.Show
End Sub

Private Sub CustMenu_Click()
CustomerForm.Show
End Sub

Private Sub CustomerListMenu_Click()
CustomerListReport.Show
End Sub

Private Sub DelRecListMenu_Click()
BookingReportForm.Show
End Sub

Private Sub DelRecMenu_Click()
ReceiptForm.Show
End Sub

Private Sub EmpListRepMenu_Click()
EmpListForm.Show
End Sub

Private Sub EmpMenu_Click()
EmpForm.Show
End Sub

Private Sub MDIForm_Load()
If Conn.State = 0 Then Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RestuarantData.mdb;Persist Security Info=False"

End Sub

Private Sub PhonebookMenu_Click()
PhoneBookingForm.Show
End Sub

Private Sub ProductMenu_Click()
ProductForm.Show
End Sub
