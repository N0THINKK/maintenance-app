VERSION 5.00
Begin VB.Form Mikrometer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patroli Harian MikroMeter"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Width           =   10095
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
      TabIndex        =   35
      Top             =   7320
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
      Left            =   9120
      TabIndex        =   34
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "No.5  Baut Pengunci tidak longgar/Dol (Visual cek, Lihat tanda pada Screw)"
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
      TabIndex        =   30
      Top             =   3840
      Width           =   10095
      Begin VB.OptionButton Option15 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option14 
         Caption         =   "NG"
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
         Left            =   4320
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Tidak ada/ Tidak Pakai"
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
         Left            =   7680
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "No.4  Kondisi Thimble, Anvil dan Spindle OK (Visual dan sentuh)"
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
      TabIndex        =   26
      Top             =   3240
      Width           =   10095
      Begin VB.OptionButton Option12 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option11 
         Caption         =   "NG"
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
         Left            =   4320
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Tidak ada/ Tidak Pakai"
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
         Left            =   7680
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "No.3  Zero setting OK (Visual cek, Layar menunjukkan ""0,000"")"
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
      TabIndex        =   22
      Top             =   2640
      Width           =   10095
      Begin VB.OptionButton Option9 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "NG"
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
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Tidak ada/ Tidak Pakai"
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
         Left            =   7680
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "No.2  Angka terbaca dengan jelas ( Visual cek, Tidak muncul huruf  ""B"", ""H"", ""INS"", atau""P"")"
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
      TabIndex        =   18
      Top             =   2040
      Width           =   10095
      Begin VB.OptionButton Option6 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "NG"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tidak ada/ Tidak Pakai"
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
         Left            =   7680
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "No.1   Ada Nomer Registrasi dan tidak Expired ( Visual Cek )"
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
      Top             =   1440
      Width           =   10095
      Begin VB.OptionButton Option3 
         Caption         =   "Tidak ada/ Tidak Pakai"
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
         Left            =   7680
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NG"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
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
      Left            =   2280
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
      Left            =   6240
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
      Left            =   8640
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
      Left            =   9120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   6600
      Picture         =   "PHMikrometer.frx":0000
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   4080
      Picture         =   "PHMikrometer.frx":947A
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   1560
      Picture         =   "PHMikrometer.frx":138B4
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   8040
      Picture         =   "PHMikrometer.frx":1D11A
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   5640
      Picture         =   "PHMikrometer.frx":2639C
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   1515
      Left            =   3120
      Picture         =   "PHMikrometer.frx":2FB72
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   360
      Picture         =   "PHMikrometer.frx":54F7C
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2535
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
      TabIndex        =   37
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "PATROLI HARIAN MIKROMETER"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   4815
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
      TabIndex        =   11
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
      Left            =   2400
      TabIndex        =   10
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
      Left            =   6360
      TabIndex        =   9
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
      Left            =   8640
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "JAM"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Mikrometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public con As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Private Sub Command1_Click()

Dim no1
Dim no2
Dim no3
Dim no4
Dim no5


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

If Option1 = True Then
no1 = "O"
Else
If Option2 = True Then
no1 = "NG"
Else
If Option3 = True Then
no1 = "N/A"
End If
End If
End If

If Option6 = True Then
no2 = "O"
Else
If Option5 = True Then
no2 = "NG"
Else
If Option4 = True Then
no2 = "N/A"
End If
End If
End If

If Option9 = True Then
no3 = "O"
Else
If Option8 = True Then
no3 = "NG"
Else
If Option7 = True Then
no3 = "N/A"
End If
End If
End If

If Option12 = True Then
no4 = "O"
Else
If Option11 = True Then
no4 = "NG"
Else
If Option10 = True Then
no4 = "N/A"
End If
End If
End If

If Option15 = True Then
no5 = "O"
Else
If Option14 = True Then
no5 = "NG"
Else
If Option13 = True Then
no5 = "N/A"
End If
End If
End If

Set rs1 = New ADODB.Recordset
rs1.Open "select*from Mikro", con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Mikro (Tahun,Bulan,Tanggal,Jam,Shift,NIK,NoMesin,AdaNomerRegistrasidantidakExpired,Angkaterbacadenganjelas,ZerosettingOK,KondisiThimbleAnvildanSpindleOK,BautPenguncitidaklonggarAtauDol,Keterangan) values ('" & Label9 & "','" & Label8 & "','" & Text1 & "','" & Label6 & "','" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & no1 & "','" & no2 & "','" & no3 & "','" & no4 & "','" & no5 & "','" & Text2 & "')"

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

End Sub

Private Sub Command2_Click()
DataMikrometer.Show
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

Call Mkr

End Sub

Private Sub Timer1_Timer()
Label6 = time
End Sub

Private Sub Mkr()

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\PHMikrometer.mdb"
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Mikro", con, adOpenDynamic, adLockOptimistic
'Set DataGrid2.DataSource = rs2
'rs2.Sort = DataGrid2.Columns(36).DataField & " DESC"
End Sub
