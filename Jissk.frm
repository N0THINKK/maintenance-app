VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form Jissk 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   7800
      TabIndex        =   21
      Text            =   "0"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Text            =   "0"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Text            =   "0"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Text            =   "0"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Text            =   "0"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Text            =   "0"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INPUT"
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
      TabIndex        =   13
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
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
      TabIndex        =   11
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   4935
      Left            =   5520
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "HASIL PENGUKURAN  MD12"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Acc sisi B"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Terminal sisi B"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Acc sisi A"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Terminal sisi A"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Jissk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public jon As New ADODB.Connection
Public rs8 As New ADODB.Recordset

Public Bon As New ADODB.Connection
Public rs3 As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into LKO (Tanggal,NoMesin,Shift,NIK,Sequen,Waktu,CutL,LotIDWire,Kombinasi,TermA,SealA,TermB,SealB,LotIDTermA,FCHTermA,FCWTermA,RCHTermA,RCWTermA,LotIDTermB,FCHTermB,FCWTermB,RCHTermB,RCWTermB,KodeDefect,QtyDefect,QtyProduct,No4M) values ('" _
& Text11 & "','" & Combo1 & "','" & Combo2 & "','" & Combo4 & "','" & Text1 & "','" & Label30 & "','" & Text20 & "','" & Text2 & "','" & Label25 & "','" & Text18 & "','" & Label5 & "','" & Text19 & "','" & Label6 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "','" & Text8 & "','" & Text9 & "','" & Text10 & "','" & Text14 & "','" & Text15 & "','" & Combo3 & "','" & Text16 & "','" & Text17 & "','" & Text21 & "')"

Con.Execute SQLTambah

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

FrmOP.Show
End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(11)
Text2.Text = DataGrid1.Columns(12)
'FrmOP.Text4 = DataGrid1.Columns(17)
'FrmOP.Text5 = DataGrid1.Columns(18)
'FrmOP.Text6 = DataGrid1.Columns(21)
'FrmOP.Text7 = DataGrid1.Columns(22)
Text5.Text = DataGrid1.Columns(17)
Text6.Text = DataGrid1.Columns(18)
Text7.Text = DataGrid1.Columns(21)
Text8.Text = DataGrid1.Columns(22)

End Sub

Private Sub DataGrid2_Click()
Text3.Text = DataGrid2.Columns(14)
Text4.Text = DataGrid2.Columns(15)
'FrmOP.Text9 = DataGrid2.Columns(25)
'FrmOP.Text10 = DataGrid2.Columns(26)
'FrmOP.Text14 = DataGrid2.Columns(29)
'FrmOP.Text15 = DataGrid2.Columns(30)
Text9.Text = DataGrid1.Columns(25)
Text10.Text = DataGrid1.Columns(26)
Text11.Text = DataGrid1.Columns(29)
Text12.Text = DataGrid1.Columns(30)

End Sub

Private Sub Form_Activate()
Set jon = New ADODB.Connection
Set rs8 = New ADODB.Recordset
jon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Jissk.mdb"

Set Bon = New ADODB.Connection
Set rs3 = New ADODB.Recordset
Bon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DBchcw.mdb"

End Sub


Private Sub Form_Load()
Set jon = New ADODB.Connection
jon.CursorLocation = adUseClient
jon.Provider = "Microsoft.jet.oledb.4.0"
jon.Open App.Path & "\Jissk.mdb"
Call MD12
rs8.Sort = DataGrid1.Columns(36).DataField & " DESC"

Set Bon = New ADODB.Connection
Bon.CursorLocation = adUseClient
Bon.Provider = "Microsoft.jet.oledb.4.0"
Bon.Open App.Path & "\DBchcw.mdb"
'Call CHCW

End Sub

Private Sub MD12()
Set rs8 = New ADODB.Recordset
rs8.Open "select*from Hasil", jon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs8
With DataGrid1
.Columns(0).Visible = False
.Columns(1).Visible = False
.Columns(2).Visible = True
.Columns(3).Visible = False
.Columns(4).Visible = True
.Columns(5).Visible = False
.Columns(6).Visible = False
.Columns(7).Visible = False
.Columns(8).Visible = False
.Columns(9).Visible = False
.Columns(10).Visible = False
.Columns(11).Visible = True
.Columns(12).Visible = True
.Columns(13).Visible = False
.Columns(14).Visible = False
.Columns(15).Visible = False
.Columns(16).Visible = False
.Columns(17).Visible = True
.Columns(18).Visible = True
.Columns(19).Visible = False
.Columns(20).Visible = False
.Columns(21).Visible = True
.Columns(22).Visible = True
.Columns(23).Visible = False
.Columns(24).Visible = False
.Columns(25).Visible = False
.Columns(26).Visible = False
.Columns(27).Visible = False
.Columns(28).Visible = False
.Columns(29).Visible = False
.Columns(30).Visible = False
.Columns(31).Visible = False
.Columns(32).Visible = False
.Columns(33).Visible = False
.Columns(34).Visible = False
.Columns(35).Visible = False
.Columns(36).Visible = True
.Columns(37).Visible = False
.Columns(38).Visible = False
.Columns(39).Visible = True
.Columns(40).Visible = True
.Columns(41).Visible = False
.Columns(42).Visible = False
.Columns(43).Visible = False
.Columns(44).Visible = False
.Columns(45).Visible = False
.Columns(46).Visible = False
End With

Set DataGrid2.DataSource = rs8
With DataGrid2
.Columns(0).Visible = False
.Columns(1).Visible = False
.Columns(2).Visible = True
.Columns(3).Visible = False
.Columns(4).Visible = True
.Columns(5).Visible = False
.Columns(6).Visible = False
.Columns(7).Visible = False
.Columns(8).Visible = False
.Columns(9).Visible = False
.Columns(10).Visible = False
.Columns(11).Visible = False
.Columns(12).Visible = False
.Columns(13).Visible = False
.Columns(14).Visible = True
.Columns(15).Visible = True
.Columns(16).Visible = False
.Columns(17).Visible = False
.Columns(18).Visible = False
.Columns(19).Visible = False
.Columns(20).Visible = False
.Columns(21).Visible = False
.Columns(22).Visible = False
.Columns(23).Visible = False
.Columns(24).Visible = False
.Columns(25).Visible = True
.Columns(26).Visible = True
.Columns(27).Visible = False
.Columns(28).Visible = False
.Columns(29).Visible = True
.Columns(30).Visible = True
.Columns(31).Visible = False
.Columns(32).Visible = False
.Columns(33).Visible = False
.Columns(34).Visible = False
.Columns(35).Visible = False
.Columns(36).Visible = True
.Columns(37).Visible = False
.Columns(38).Visible = False
.Columns(39).Visible = True
.Columns(40).Visible = True
.Columns(41).Visible = False
.Columns(42).Visible = False
.Columns(43).Visible = False
.Columns(44).Visible = False
.Columns(45).Visible = False
.Columns(46).Visible = False

End With

End Sub

