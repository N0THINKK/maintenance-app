VERSION 5.00
Begin VB.Form PhAplC 
   Caption         =   "Patroli Harian Aplikator"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   405
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
   ScaleHeight     =   9480
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   1320
      TabIndex        =   68
      Top             =   9000
      Width           =   7575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   9240
      TabIndex        =   66
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Record"
      Height          =   675
      Left            =   9240
      TabIndex        =   65
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   615
      Left            =   9240
      TabIndex        =   64
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Frame Frame12 
      Caption         =   "Applikator 12"
      Height          =   495
      Left            =   6360
      TabIndex        =   60
      Top             =   8400
      Width           =   2535
      Begin VB.OptionButton Option24 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option23 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Applikator 11"
      Height          =   495
      Left            =   6360
      TabIndex        =   57
      Top             =   7800
      Width           =   2535
      Begin VB.OptionButton Option22 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option21 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Applikator 10"
      Height          =   495
      Left            =   6360
      TabIndex        =   54
      Top             =   7200
      Width           =   2535
      Begin VB.OptionButton Option20 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option19 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Applikator 9"
      Height          =   495
      Left            =   6360
      TabIndex        =   51
      Top             =   6600
      Width           =   2535
      Begin VB.OptionButton Option18 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option17 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Applikator 8"
      Height          =   495
      Left            =   6360
      TabIndex        =   48
      Top             =   6000
      Width           =   2535
      Begin VB.OptionButton Option16 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option15 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Applikator 7"
      Height          =   495
      Left            =   6360
      TabIndex        =   45
      Top             =   5400
      Width           =   2535
      Begin VB.OptionButton Option14 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option13 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Applikator 6"
      Height          =   495
      Left            =   6360
      TabIndex        =   42
      Top             =   4800
      Width           =   2535
      Begin VB.OptionButton Option12 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option11 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Applikator 5"
      Height          =   495
      Left            =   6360
      TabIndex        =   39
      Top             =   4200
      Width           =   2535
      Begin VB.OptionButton Option10 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option9 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Applikator 4"
      Height          =   495
      Left            =   6360
      TabIndex        =   36
      Top             =   3600
      Width           =   2535
      Begin VB.OptionButton Option8 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Applikator 3"
      Height          =   495
      Left            =   6360
      TabIndex        =   33
      Top             =   3000
      Width           =   2535
      Begin VB.OptionButton Option6 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Applikator 2"
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   2400
      Width           =   2535
      Begin VB.OptionButton Option4 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applikator 1"
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   1800
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "NG"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sisi A"
      Height          =   615
      Left            =   9240
      TabIndex        =   26
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8760
      TabIndex        =   6
      Top             =   960
      Width           =   1575
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
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   9360
      TabIndex        =   67
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   120
      TabIndex        =   69
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
      TabIndex        =   63
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Sisi B"
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
      TabIndex        =   25
      Top             =   840
      Width           =   1095
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
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7620
      Left            =   120
      Picture         =   "PhAppl sisi C.frx":0000
      Top             =   1320
      Width           =   3675
   End
   Begin VB.Label Label8 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "JAM"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "No. Mesin"
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "NIK"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
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
Attribute VB_Name = "PhAplC"
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

Private Sub Command1_Click()
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

If Combo3 = "" Then
MsgBox " No.Mesin belum terisi"
Exit Sub
End If

If Option1 = False And Option2 = False Then
MsgBox " Applikator 1 belum terisi"
Exit Sub
End If

If Option1 = True Then
Appl1 = "O"
Else
If Option2 = True Then
Appl1 = "NG"
End If
End If

If Option4 = True Then
Appl2 = "O"
Else
If Option3 = True Then
Appl2 = "NG"
End If
End If

If Option5 = True Then
Appl3 = "O"
Else
If Option6 = True Then
Appl3 = "NG"
End If
End If

If Option7 = True Then
Appl4 = "O"
Else
If Option8 = True Then
Appl4 = "NG"
End If
End If

If Option9 = True Then
Appl5 = "O"
Else
If Option10 = True Then
Appl5 = "NG"
End If
End If

If Option12 = True Then
Appl6 = "O"
Else
If Option11 = True Then
Appl6 = "NG"
End If
End If

If Option13 = True Then
Appl7 = "O"
Else
If Option14 = True Then
Appl7 = "NG"
End If
End If

If Option16 = True Then
Appl8 = "O"
Else
If Option15 = True Then
Appl8 = "NG"
End If
End If

If Option18 = True Then
Appl9 = "O"
Else
If Option17 = True Then
Appl9 = "NG"
End If
End If

If Option20 = True Then
Appl10 = "O"
Else
If Option19 = True Then
Appl10 = "NG"
End If
End If

If Option21 = True Then
Appl11 = "O"
Else
If Option22 = True Then
Appl11 = "NG"
End If
End If

If Option23 = True Then
Appl12 = "O"
Else
If Option24 = True Then
Appl12 = "NG"
End If
End If

Set rs3 = New ADODB.Recordset
rs3.Open "select*from Appl_B", con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Appl_B (Tahun,Bulan,Tanggal,Shift,NIK,NoMesin,Aplikator1,KondisiAplikator1,Aplikator2,KondisiAplikator2,Aplikator3,KondisiAplikator3,Aplikator4,KondisiAplikator4,Aplikator5,KondisiAplikator5,Aplikator6,KondisiAplikator6,Aplikator7,KondisiAplikator7,Aplikator8,KondisiAplikator8,Aplikator9,KondisiAplikator9,Aplikator10,KondisiAplikator10,Aplikator11,KondisiAplikator11,Aplikator12,KondisiAplikator12,Jam,Keterangan) values ('" _
& Label7 & "','" & Label8 & "','" & Text1 & "','" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & Text2 & "','" & Appl1 & "','" & Text3 & "','" & Appl2 & "','" & Text4 & "','" & Appl3 & "','" & Text5 & "','" & Appl4 & "','" & Text6 & "','" & Appl5 & "','" & Text7 & "','" & Appl6 & "','" & Text8 & "','" & Appl7 & "','" & Text9 & "','" & Appl8 & "','" & Text10 & "','" & Appl9 & "','" & Text11 & "','" & Appl10 & "','" & Text12 & "','" & Appl11 & "','" & Text13 & "','" & Appl12 & "','" & Label6 & "','" & Text14 & "')"

con.Execute SQLTambah

  SetTimer hWnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hWnd, "Data berhasil disimpan", _
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


End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1 = Format(Now, "yyyy/mm/dd")
Label6 = Format(Now, "hh:mm:ss")
Label7 = Format(Now, "yyyy")
Label8 = Format(Now, "mm")

Combo3.Text = NoMesin.Text21.Text

With Combo1
.AddItem "A"
.AddItem "B"
.AddItem "NS"
End With

End Sub

Private Sub Form_Load()

Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\cfm.mdb"

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\DbPHAppl.mdb"

Call AplikatorB

Text2.Text = List1.list(0)
Text3.Text = List1.list(1)
Text4.Text = List1.list(2)
Text5.Text = List1.list(3)
Text6.Text = List1.list(4)
Text7.Text = List1.list(5)
Text8.Text = List1.list(6)
Text9.Text = List1.list(7)
Text10.Text = List1.list(8)
Text11.Text = List1.list(9)
Text12.Text = List1.list(10)
Text13.Text = List1.list(11)

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label22_Click()

End Sub

Private Sub AplikatorB()

Dim Ldata As ListItem

Set rs1 = New ADODB.Recordset
rs1.Open "select*from SisiB", kon, adOpenDynamic, adLockOptimistic


Set rs2 = New ADODB.Recordset
rs2.Open "select*from CfmB", kon
List1.Clear
Do While Not rs2.EOF
List1.AddItem rs2!No & ". " & rs2!Aplikator
rs2.MoveNext
Loop
rs2.Close

End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
Text12.Text = ListView1
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Text1_Click()
Text1 = Format(Now, "yyyy/mm/dd")
End Sub
