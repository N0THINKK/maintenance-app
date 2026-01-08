VERSION 5.00
Begin VB.Form PhAplA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patroli Harian Aplikator"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9360
      TabIndex        =   95
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8760
      TabIndex        =   94
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Check12 
      Height          =   255
      Left            =   3960
      TabIndex        =   93
      Top             =   8640
      Width           =   255
   End
   Begin VB.CheckBox Check11 
      Height          =   255
      Left            =   3960
      TabIndex        =   92
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check10 
      Height          =   255
      Left            =   3960
      TabIndex        =   91
      Top             =   7440
      Width           =   255
   End
   Begin VB.CheckBox Check9 
      Height          =   255
      Left            =   3960
      TabIndex        =   90
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox Check8 
      Height          =   255
      Left            =   3960
      TabIndex        =   89
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox Check7 
      Height          =   255
      Left            =   3960
      TabIndex        =   88
      Top             =   5640
      Width           =   255
   End
   Begin VB.CheckBox Check6 
      Height          =   255
      Left            =   3960
      TabIndex        =   87
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check5 
      Height          =   255
      Left            =   3960
      TabIndex        =   86
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Height          =   255
      Left            =   3960
      TabIndex        =   85
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Height          =   255
      Left            =   3960
      TabIndex        =   84
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   3960
      TabIndex        =   83
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   3960
      TabIndex        =   82
      Top             =   2040
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   9480
      TabIndex        =   69
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   1320
      TabIndex        =   66
      Top             =   9000
      Width           =   7575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   9240
      TabIndex        =   65
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Record"
      Height          =   615
      Left            =   9360
      TabIndex        =   64
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   615
      Left            =   9360
      TabIndex        =   63
      Top             =   8280
      Width           =   975
   End
   Begin VB.Frame Frame12 
      Caption         =   "Applikator 12"
      Height          =   495
      Left            =   6360
      TabIndex        =   59
      Top             =   8400
      Width           =   2775
      Begin VB.OptionButton Option36 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   81
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option24 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option23 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Applikator 11"
      Height          =   495
      Left            =   6360
      TabIndex        =   56
      Top             =   7800
      Width           =   2775
      Begin VB.OptionButton Option35 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option22 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option21 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Applikator 10"
      Height          =   495
      Left            =   6360
      TabIndex        =   53
      Top             =   7200
      Width           =   2775
      Begin VB.OptionButton Option34 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option20 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option19 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Applikator 9"
      Height          =   495
      Left            =   6360
      TabIndex        =   50
      Top             =   6600
      Width           =   2775
      Begin VB.OptionButton Option33 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option18 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option17 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Applikator 8"
      Height          =   495
      Left            =   6360
      TabIndex        =   47
      Top             =   6000
      Width           =   2775
      Begin VB.OptionButton Option32 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   77
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option16 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option15 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Applikator 7"
      Height          =   495
      Left            =   6360
      TabIndex        =   44
      Top             =   5400
      Width           =   2775
      Begin VB.OptionButton Option31 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   76
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option14 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option13 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Applikator 6"
      Height          =   495
      Left            =   6360
      TabIndex        =   41
      Top             =   4800
      Width           =   2775
      Begin VB.OptionButton Option30 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   75
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option12 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option11 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Applikator 5"
      Height          =   495
      Left            =   6360
      TabIndex        =   38
      Top             =   4200
      Width           =   2775
      Begin VB.OptionButton Option29 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option10 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option9 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Applikator 4"
      Height          =   495
      Left            =   6360
      TabIndex        =   35
      Top             =   3600
      Width           =   2775
      Begin VB.OptionButton Option28 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option8 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Applikator 3"
      Height          =   495
      Left            =   6360
      TabIndex        =   32
      Top             =   3000
      Width           =   2775
      Begin VB.OptionButton Option27 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   72
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Applikator 2"
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   2400
      Width           =   2775
      Begin VB.OptionButton Option26 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   71
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applikator 1"
      Height          =   495
      Left            =   6360
      TabIndex        =   26
      Top             =   1800
      Width           =   2775
      Begin VB.OptionButton Option25 
         Caption         =   "N/A"
         Height          =   255
         Left            =   2040
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NG"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sisi B"
      Height          =   615
      Left            =   9360
      TabIndex        =   25
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   68
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kondisi Aplikator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   62
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Sisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      Caption         =   "No. Aplikator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7620
      Left            =   120
      Picture         =   "PhAppl sisi A.frx":0000
      Top             =   1320
      Width           =   3675
   End
   Begin VB.Label Label8 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "JAM"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "No. Mesin"
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "NIK"
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Shift"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "PATROLI HARIAN APLIKATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "PhAplA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset

Public con As New ADODB.Connection
Public rs3 As New ADODB.Recordset

Option Explicit

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.Enabled = True
Text2.Text = List1.List(0)
Frame1.Enabled = True
ElseIf Check1.Value = 0 Then
Text2.Enabled = False
Text2.Text = ""
Frame1.Enabled = False
End If

End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then
Text11.Enabled = True
Text11.Text = List1.List(9)
Frame1.Enabled = True
ElseIf Check10.Value = 0 Then
Text11.Enabled = False
Text11.Text = ""
Frame10.Enabled = False
End If
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then
Text12.Enabled = True
Text12.Text = List1.List(10)
Frame11.Enabled = True
ElseIf Check11.Value = 0 Then
Text12.Enabled = False
Text12.Text = ""
Frame11.Enabled = False
End If
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then
Text13.Enabled = True
Text13.Text = List1.List(11)
Frame12.Enabled = True
ElseIf Check12.Value = 0 Then
Text13.Enabled = False
Text13.Text = ""
Frame12.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text3.Enabled = True
Text3.Text = List1.List(1)
Frame2.Enabled = True
ElseIf Check2.Value = 0 Then
Text3.Enabled = False
Text3.Text = ""
Frame2.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Text4.Enabled = True
Text4.Text = List1.List(2)
Frame3.Enabled = True
ElseIf Check3.Value = 0 Then
Text4.Enabled = False
Text4.Text = ""
Frame3.Enabled = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Text5.Enabled = True
Text5.Text = List1.List(3)
Frame4.Enabled = True
ElseIf Check4.Value = 0 Then
Text5.Enabled = False
Text5.Text = ""
Frame4.Enabled = False
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
Text6.Enabled = True
Text6.Text = List1.List(4)
Frame5.Enabled = True
ElseIf Check5.Value = 0 Then
Text6.Enabled = False
Text6.Text = ""
Frame5.Enabled = False
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
Text7.Enabled = True
Text7.Text = List1.List(5)
Frame6.Enabled = True
ElseIf Check6.Value = 0 Then
Text7.Enabled = False
Text7.Text = ""
Frame6.Enabled = False
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
Text8.Enabled = True
Text8.Text = List1.List(6)
Frame7.Enabled = True
ElseIf Check7.Value = 0 Then
Text8.Enabled = False
Text8.Text = ""
Frame7.Enabled = False
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
Text9.Enabled = True
Text9.Text = List1.List(7)
Frame8.Enabled = True
ElseIf Check8.Value = 0 Then
Text9.Enabled = False
Text9.Text = ""
Frame8.Enabled = False
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
Text10.Enabled = True
Text10.Text = List1.List(8)
Frame9.Enabled = True
ElseIf Check9.Value = 0 Then
Text10.Enabled = False
Text10.Text = ""
Frame9.Enabled = False
End If
End Sub

Private Sub Command5_Click()

If Command5.Caption = "NEXT" Then
Text2.Text = List1.List(12)
Text3.Text = List1.List(13)
Text4.Text = List1.List(14)
Text5.Text = List1.List(15)
Text6.Text = List1.List(16)
Text7.Text = List1.List(17)
Text8.Text = List1.List(18)
Text9.Text = List1.List(19)
Text10.Text = List1.List(20)
Text11.Text = List1.List(21)
Text12.Text = List1.List(22)
Text13.Text = List1.List(23)

Command5.Caption = "BACK"

Else
If Command5.Caption = "BACK" Then
Text2.Text = List1.List(0)
Text3.Text = List1.List(1)
Text4.Text = List1.List(2)
Text5.Text = List1.List(3)
Text6.Text = List1.List(4)
Text7.Text = List1.List(5)
Text8.Text = List1.List(6)
Text9.Text = List1.List(7)
Text10.Text = List1.List(8)
Text11.Text = List1.List(9)
Text12.Text = List1.List(10)
Text13.Text = List1.List(11)
Command5.Caption = "NEXT"

End If
End If

End Sub

Private Sub text15_Change()
'Call AplikatorA
End Sub

Private Sub Command1_Click()
Unload Me

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "DAplikator Sisi B ditampilkan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

PhAplB.Show
End Sub

Private Sub Command2_Click()

Dim Appl1
Dim Appl2
Dim Appl3
Dim Appl4
Dim Appl5
Dim Appl6
Dim Appl7
Dim Appl8
Dim Appl9
Dim Appl10
Dim Appl11
Dim Appl12

If Text1 = "" Then
MsgBox " Data Tanggal belum terisi"
Exit Sub
End If

If Combo1 = "" Then
MsgBox " Shift Belum terisi"
Exit Sub
End If

If Combo2 = "" Then
MsgBox " NIK belum terisi"
Exit Sub
End If

If Text15 = "" Then
MsgBox " No.Mesin belum terisi"
Exit Sub
End If

If Option1 = True Then
Appl1 = "O"
Else
If Option2 = True Then
Appl1 = "NG"
Else
If Option25 = True Then
Appl1 = "N/A"
End If
End If
End If

If Option4 = True Then
Appl2 = "O"
Else
If Option3 = True Then
Appl2 = "NG"
Else
If Option26 = True Then
Appl2 = "N/A"
End If
End If
End If

If Option5 = True Then
Appl3 = "O"
Else
If Option6 = True Then
Appl3 = "NG"
Else
If Option27 = True Then
Appl3 = "N/A"
End If
End If
End If

If Option7 = True Then
Appl4 = "O"
Else
If Option8 = True Then
Appl4 = "NG"
Else
If Option28 = True Then
Appl4 = "N/A"
End If
End If
End If

If Option9 = True Then
Appl5 = "O"
Else
If Option10 = True Then
Appl5 = "NG"
Else
If Option29 = True Then
Appl5 = "N/A"
End If
End If
End If

If Option12 = True Then
Appl6 = "O"
Else
If Option11 = True Then
Appl6 = "NG"
Else
If Option30 = True Then
Appl6 = "N/A"
End If
End If
End If

If Option13 = True Then
Appl7 = "O"
Else
If Option14 = True Then
Appl7 = "NG"
Else
If Option31 = True Then
Appl7 = "N/A"
End If
End If
End If

If Option16 = True Then
Appl8 = "O"
Else
If Option15 = True Then
Appl8 = "NG"
Else
If Option32 = True Then
Appl8 = "N/A"
End If
End If
End If

If Option18 = True Then
Appl9 = "O"
Else
If Option17 = True Then
Appl9 = "NG"
Else
If Option33 = True Then
Appl9 = "N/A"
End If
End If
End If

If Option20 = True Then
Appl10 = "O"
Else
If Option19 = True Then
Appl10 = "NG"
Else
If Option34 = True Then
Appl10 = "N/A"
End If
End If
End If

If Option21 = True Then
Appl11 = "O"
Else
If Option22 = True Then
Appl11 = "NG"
Else
If Option35 = True Then
Appl11 = "N/A"
End If
End If
End If

If Option23 = True Then
Appl12 = "O"
Else
If Option24 = True Then
Appl12 = "NG"
Else
If Option36 = True Then
Appl12 = "N/A"
End If
End If
End If

Set rs3 = New ADODB.Recordset
rs3.Open "select*from Appl_A", con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Appl_B (Tahun,Bulan,Tanggal,Shift,NIK,NoMesin,Aplikator1,KondisiAplikator1,Aplikator2,KondisiAplikator2,Aplikator3,KondisiAplikator3,Aplikator4,KondisiAplikator4,Aplikator5,KondisiAplikator5,Aplikator6,KondisiAplikator6,Aplikator7,KondisiAplikator7,Aplikator8,KondisiAplikator8,Aplikator9,KondisiAplikator9,Aplikator10,KondisiAplikator10,Aplikator11,KondisiAplikator11,Aplikator12,KondisiAplikator12,Jam,Keterangan) values ('" _
& Label7 & "','" & Label8 & "','" & Text1 & "','" & Combo1 & "','" & Combo2 & "','" & Text15 & "','" & Text2 & "','" & Appl1 & "','" & Text3 & "','" & Appl2 & "','" & Text4 & "','" & Appl3 & "','" & Text5 & "','" & Appl4 & "','" & Text6 & "','" & Appl5 & "','" & Text7 & "','" & Appl6 & "','" & Text8 & "','" & Appl7 & "','" & Text9 & "','" & Appl8 & "','" & Text10 & "','" & Appl9 & "','" & Text11 & "','" & Appl10 & "','" & Text12 & "','" & Appl11 & "','" & Text13 & "','" & Appl12 & "','" & Label6 & "','" & Text14 & "')"

con.Execute SQLTambah

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
      
      
Option1 = False
Option2 = False
Option3 = False
Option4 = False
Option5 = False
Option6 = False
Option7 = False
Option8 = False
Option9 = False
Option10 = False
Option11 = False
Option12 = False
Option13 = False
Option14 = False
Option15 = False
Option16 = False
Option17 = False
Option18 = False
Option19 = False
Option20 = False
Option21 = False
Option22 = False
Option23 = False
Option24 = False

Unload Me
PhAplB.Show
PhAplB.Combo1 = DataPHApplA.DataGrid1.Columns(3)
PhAplB.Combo2 = DataPHApplA.DataGrid1.Columns(4)

End Sub

Private Sub Command3_Click()
DataPHApplB.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1 = Format(Now, "yyyy/mm/dd")
Label6 = Format(Now, "hh:mm:ss")
Label7 = Format(Now, "yyyy")
Label8 = Format(Now, "mm")



End Sub

Private Sub Form_Load()

Text15.Text = NoMesin.Text21

Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\MasterAplikator.mdb"
Call AplikatorA

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\DbPHAppl.mdb"

Text2.Text = List1.List(0)
Text3.Text = List1.List(1)
Text4.Text = List1.List(2)
Text5.Text = List1.List(3)
Text6.Text = List1.List(4)
Text7.Text = List1.List(5)
Text8.Text = List1.List(6)
Text9.Text = List1.List(7)
Text10.Text = List1.List(8)
Text11.Text = List1.List(9)
Text12.Text = List1.List(10)
Text13.Text = List1.List(11)

'Combo1.Text = DataPHApplA.DataGrid1.Columns(3)
'Combo2.Text = DataPHApplA.DataGrid1.Columns(4)

With Combo1
.AddItem "A"
.AddItem "B"
.AddItem "NS"
End With

Frame1.Caption = Text2.Text
Frame2.Caption = Text3.Text
Frame3.Caption = Text4.Text
Frame4.Caption = Text5.Text
Frame5.Caption = Text6.Text
Frame6.Caption = Text7.Text
Frame7.Caption = Text8.Text
Frame8.Caption = Text9.Text
Frame9.Caption = Text10.Text
Frame10.Caption = Text11.Text
Frame11.Caption = Text12.Text
Frame12.Caption = Text13.Text

Call kriteria1

Call kriteria2

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label22_Click()

End Sub


Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
'Text12.Text = ListView1
End Sub

Private Sub Text1_Click()
Text1 = Format(Now, "yyyy/mm/dd")
End Sub
Private Sub AplikatorA()

Form_Activate
rs2.Open "select * from Appl where NoMesin = '" & Text15 & "' and Sisi = '" & Label13 & "'", kon
If Not rs2.EOF Then
List1.Clear
Do While Not rs2.EOF
List1.AddItem rs2!NoApplikator & " " & rs2!Type & " " & rs2!Serial
rs2.MoveNext
Loop

Else
MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi"

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data yang anda cari tidak ada", _
   MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

End If

rs2.Close
End Sub

Private Sub kriteria1()

If Text2 <> "" Then
Check1.Value = 1
End If

If Text3 <> "" Then
Check2.Value = 1
End If

If Text4 <> "" Then
Check3.Value = 1
End If

If Text5 <> "" Then
Check4.Value = 1
End If

If Text6 <> "" Then
Check5.Value = 1
End If

If Text7 <> "" Then
Check6.Value = 1
End If

If Text8 <> "" Then
Check7.Value = 1
End If

If Text9 <> "" Then
Check8.Value = 1
End If
If Text10 <> "" Then
Check9.Value = 1
End If

If Text11 <> "" Then
Check10.Value = 1
End If

If Text12 <> "" Then
Check11.Value = 1
End If

If Text13 <> "" Then
Check12.Value = 1
End If


End Sub


Private Sub kriteria2()

If Check1.Value = 1 Then
Text2.Enabled = True
Text2.Text = List1.List(0)
Frame1.Enabled = True
ElseIf Check1.Value = 0 Then
Text2.Enabled = False
Text2.Text = ""
Frame1.Enabled = False
End If

If Check2.Value = 1 Then
Text3.Enabled = True
Text3.Text = List1.List(1)
Frame2.Enabled = True
ElseIf Check2.Value = 0 Then
Text3.Enabled = False
Text3.Text = ""
Frame2.Enabled = False
End If

If Check3.Value = 1 Then
Text4.Enabled = True
Text4.Text = List1.List(2)
Frame3.Enabled = True
ElseIf Check3.Value = 0 Then
Text4.Enabled = False
Text4.Text = ""
Frame3.Enabled = False
End If

If Check4.Value = 1 Then
Text5.Enabled = True
Text5.Text = List1.List(3)
Frame4.Enabled = True
ElseIf Check4.Value = 0 Then
Text5.Enabled = False
Text5.Text = ""
Frame4.Enabled = False
End If

If Check5.Value = 1 Then
Text6.Enabled = True
Text6.Text = List1.List(4)
Frame5.Enabled = True
ElseIf Check5.Value = 0 Then
Text6.Enabled = False
Text6.Text = ""
Frame5.Enabled = False
End If

If Check6.Value = 1 Then
Text7.Enabled = True
Text7.Text = List1.List(5)
Frame6.Enabled = True
ElseIf Check6.Value = 0 Then
Text7.Enabled = False
Text7.Text = ""
Frame6.Enabled = False
End If

If Check7.Value = 1 Then
Text8.Enabled = True
Text8.Text = List1.List(6)
Frame7.Enabled = True
ElseIf Check7.Value = 0 Then
Text8.Enabled = False
Text8.Text = ""
Frame7.Enabled = False
End If

If Check8.Value = 1 Then
Text9.Enabled = True
Text9.Text = List1.List(7)
Frame8.Enabled = True
ElseIf Check8.Value = 0 Then
Text9.Enabled = False
Text9.Text = ""
Frame8.Enabled = False
End If

If Check9.Value = 1 Then
Text10.Enabled = True
Text10.Text = List1.List(8)
Frame9.Enabled = True
ElseIf Check9.Value = 0 Then
Text10.Enabled = False
Text10.Text = ""
Frame9.Enabled = False
End If

If Check10.Value = 1 Then
Text11.Enabled = True
Text11.Text = List1.List(9)
Frame1.Enabled = True
ElseIf Check10.Value = 0 Then
Text11.Enabled = False
Text11.Text = ""
Frame10.Enabled = False
End If

If Check11.Value = 1 Then
Text12.Enabled = True
Text12.Text = List1.List(10)
Frame11.Enabled = True
ElseIf Check11.Value = 0 Then
Text12.Enabled = False
Text12.Text = ""
Frame11.Enabled = False
End If

If Check12.Value = 1 Then
Text13.Enabled = True
Text13.Text = List1.List(11)
Frame12.Enabled = True
ElseIf Check12.Value = 0 Then
Text13.Enabled = False
Text13.Text = ""
Frame12.Enabled = False
End If

End Sub
