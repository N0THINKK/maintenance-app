VERSION 5.00
Begin VB.Form TrackDefect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracking Defect "
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4920
      TabIndex        =   33
      Text            =   "Operator"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   30
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6360
      TabIndex        =   29
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   25
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Pilih Sequen di tabel ""Barcode Kanban"" untuk menampilkan Data"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   34
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label7 
      Caption         =   "Penemu Defect"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "TRACING DEFECT PREASSY"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label17 
      Caption         =   "No. Seal"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   31
      Top             =   2400
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      X1              =   1920
      X2              =   6840
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label15 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Picture         =   "TrackingDefect.frx":0000
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label Label14 
      Caption         =   "No. Terminal"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Sequen"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "QTY Defect"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   19
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Jenis Defect"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "NIK"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "No. Mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "JAM"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "TrackingDefect.frx":046C
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   1560
      X2              =   7200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label25 
      Caption         =   "Kombinasi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label16 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "TrackDefect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public con As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Private Sub Command1_Click()


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

If Combo3 = "" Then
MsgBox " No.Mesin belum terisi"
Exit Sub
End If


Set rs1 = New ADODB.Recordset
rs1.Open "select*from Defect", con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Defect (Tahun,Bulan,Tanggal,Jam,Shift,NoMesin,NIK,Sequen,PenemuDefect,JenisDefect,Qty,NoTerminal,NoSeal,Keterangan) values ('" & Label9 & "','" & Label8 & "','" & Text1 & "','" & Label6 & "','" & Combo1 & "','" & Combo3 & "','" & Combo2 & "','" & Text5 & "','" & Combo5 & "','" & Combo4 & "','" & Text3 & "','" & Text4 & "','" & Text6 & "','" & Text2 & "')"

con.Execute SQLTambah

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Text2 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text3 = ""
Combo4 = ""
Combo5 = ""

End Sub

Private Sub Command2_Click()
DataDefect.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Text1 = Format(Now, "yyyy/mm/dd")
'Label6 = Format(Now, "hh:mm:ss")
Label9 = Format(Now, "yyyy")
Label8 = Format(Now, "mm")

Combo3.Text = NoMesin.Text21.Text

End Sub

Private Sub Form_Load()

With Combo1
.AddItem "A"
.AddItem "B"
.AddItem "NS"
End With

With Combo5
.AddItem "Operator"
.AddItem "Inspektor"
End With

With Combo4
.AddItem "A.1 Core Terurai"
.AddItem "A.2 Core Terpotong"
.AddItem "A.3 Core Rusak"
.AddItem "A.4 Core Tidak Teratur"
.AddItem "A.5 Core Maju"
.AddItem "A.6 Core Terurai"
.AddItem "A.7 Core Terpotong"
.AddItem "A.8 Core Rusak"
.AddItem "B.1 Terminal Tergores"
.AddItem "B.2 Terminal Bengkok ke atas"
.AddItem "B.3 Terminal Bengkok ke bawah"
.AddItem "B.4 Terminal Melintir"
.AddItem "B.5 Terminal Ujung Terpotong"
.AddItem "B.6 Terminal Ujung Terbuka"
.AddItem "B.7 Terminal Ujung Rusak"
.AddItem "B.8 Terminal Bridge terlalu panjang"
.AddItem "B.9 Terminal Rusak"
.AddItem "B.10 Terminal Lepas dari Circuit"
.AddItem "C.1 Front C/H terlalu tinggi"
.AddItem "C.2 Front C/H terlalu rendah"
.AddItem "C.3 Front C/W terlalu tinggi"
.AddItem "C.4 Front C/W terlalu rendah"
.AddItem "C.5 Front Flash"
.AddItem "D.1 Rear C/H terlalu tinggi"
.AddItem "D.2 Rear C/H terlalu rendah"
.AddItem "D.3 Rear C/W terlalu tinggi"
.AddItem "D.4 Rear C/W terlalu rendah"
.AddItem "D.5 Rear ada di dalam Insulasi"
.AddItem "D.6 Rear Tidak Tercrimping"
.AddItem "D.7 Rear Tidak seimbang"
.AddItem "E.1 Insulation Tercrimping"
.AddItem "E.2 Insulation Terlalu mundur"
.AddItem "E.3 Insulation Rusak"
.AddItem "E.4 Insulation Tidak rata"
.AddItem "F.1 Seal Terpotong"
.AddItem "F.2 Seal Terbalik"
.AddItem "F.3 Seal Terlalu mundur"
.AddItem "F.4 Seal Terlalu maju"
.AddItem "F.5 Seal Tercrimping"
.AddItem "F.6 Seal Tidak ada"
.AddItem "F.7 Seal Sobek"
.AddItem "G.1 Crimping Ada Benda Asing"
.AddItem "G.2 Crimping Ada 2 Terminal Tercrimping"
.AddItem "G.3 Crimping Tanpa Core"
.AddItem "G.4 Crimping Tanpa Stripping"
.AddItem "H.1 Lance Rusak"
.AddItem "H.2 Stabilizer Rusak"
.AddItem "H.3 Bellmouth Tidak Standart"
End With

Call Dft

Combo1.Text = FrmOP.Combo2
Combo2.Text = FrmOP.Combo4

End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image1_Click()
Text4.Text = Text18
Text6.Text = Label16

End Sub

Private Sub Image2_Click()
Text4.Text = Text19
Text6.Text = Label15

End Sub

Private Sub Timer1_Timer()
Label6 = time
End Sub

Private Sub Dft()

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\RecordDefect.mdb"
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Defect", con, adOpenDynamic, adLockOptimistic
'Set DataGrid2.DataSource = rs2
'rs2.Sort = DataGrid2.Columns(36).DataField & " DESC"
End Sub
