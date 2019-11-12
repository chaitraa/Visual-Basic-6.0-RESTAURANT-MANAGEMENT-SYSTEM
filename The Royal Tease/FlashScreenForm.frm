VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FlashScreenForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19020
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Snap ITC"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FlashScreenForm.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   19020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18120
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   -480
      Picture         =   "FlashScreenForm.frx":5C94F
      ScaleHeight     =   2115
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   480
      Width           =   9015
   End
   Begin VB.Timer Timer2 
      Interval        =   150
      Left            =   480
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   8400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   5640
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label per 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label complet 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   8295
   End
End
Attribute VB_Name = "FlashScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Timer1.Interval = 50
    Timer1.Enabled = True
    str = "Loading Please Wait..."
   complet.Caption = str



End Sub




Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
per.Caption = ProgressBar1.Value
If ProgressBar1.Value = 20 Then Timer1.Interval = 20
    'complet.Caption = "Loading Forms..."

If ProgressBar1.Value = 45 Then Timer1.Interval = 1
    'complet.Caption = "Loading Database Connectivity..."

If ProgressBar1.Value = 70 Then Timer1.Interval = 20
    'complet.Caption = "Loading Various Component..."

If ProgressBar1.Value = 85 Then Timer1.Interval = 60
    'complet.Caption = "Completing Please Wait..."

If ProgressBar1.Value = 100 Then
    complet.Caption = "Completed"
    
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    per.Visible = False
    complet.Visible = False
    LoginForm.Show
    Exit Sub
    End If
End Sub


Private Sub Timer2_Timer()
Picture1.Left = Picture1.Left + 1000
If Picture1.Left > 16000 Then
Picture1.Left = -7000
End If
End Sub


