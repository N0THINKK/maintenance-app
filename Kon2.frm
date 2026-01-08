VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form Kon2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Koneksi"
   ClientHeight    =   8835
   ClientLeft      =   8100
   ClientTop       =   1455
   ClientWidth     =   10875
   ControlBox      =   0   'False
   Icon            =   "Kon2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1720
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   5175
      Left            =   7440
      TabIndex        =   15
      Top             =   2520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9128
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5175
      Left            =   3360
      TabIndex        =   14
      Top             =   2520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9128
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
      Height          =   5175
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9128
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
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Record Login"
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
      Left            =   8640
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "LKO"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "ProdukMaster"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9960
      TabIndex        =   6
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Koneksi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sequen"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   705
   End
End
Attribute VB_Name = "Kon2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Public con As New ADODB.Connection
Public rs2 As New ADODB.Recordset

Public Bon As New ADODB.Connection
Public rs3 As New ADODB.Recordset

Public jon As New ADODB.Connection
Public rs4 As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
Kon2.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\PrdMaster.mdb"

Set con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

Set Bon = New ADODB.Connection
Set rs3 = New ADODB.Recordset
Bon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\RecordLogin.mdb"

Text3.Text = FrmOP.Text1
Text1.Text = FrmOP.Text26
Text4.Text = FrmOP.Text28

End Sub

Private Sub Form_Load()

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\DbLKO.mdb"
Call DbLKO

Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\PrdMaster.mdb"
Call Kanban

Set Bon = New ADODB.Connection
Bon.CursorLocation = adUseClient
Bon.Provider = "Microsoft.jet.oledb.4.0"
Bon.Open App.Path & "\RecordLogin.mdb"
Call Masuk


Set jon = New ADODB.Connection
jon.CursorLocation = adUseClient
jon.Provider = "Microsoft.jet.oledb.4.0"
jon.Open App.Path & "\Jissk.mdb"
'Call MuatUlang
End Sub

Private Sub DbLKO()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", con, adOpenDynamic, adLockOptimistic
Set DataGrid2.DataSource = rs2
rs2.Sort = DataGrid2.Columns(36).DataField & " DESC"


End Sub

Private Sub Kanban()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Prdmst", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
End Sub
Private Sub Masuk()
Set rs3 = New ADODB.Recordset
rs3.Open "select*from PIC", Bon, adOpenDynamic, adLockOptimistic
Set DataGrid3.DataSource = rs3
rs3.Sort = DataGrid3.Columns(5).DataField & " DESC"

End Sub
Private Sub MuatUlang()
Set rs4 = New ADODB.Recordset
rs4.Open "select*from Prdlog", jon, adOpenDynamic, adLockOptimistic
Set DataGrid4.DataSource = rs4
rs4.Sort = DataGrid4.Columns(3).DataField & " DESC"

End Sub

Public Sub PANGGILKANBAN2()

Form_Activate
rs1.Open "select * from Prdmst where Field1 = '" & Text2 & "'", kon
If Not rs1.EOF Then
Text23 = rs1!Field1
FrmOP.Label25.Caption = rs1!Field44
FrmOP.Text18.Text = rs1!Field12
FrmOP.Text19.Text = rs1!Field13
FrmOP.Text20.Text = rs1!Field6
FrmOP.Label39.Caption = rs1!Field14
FrmOP.Label40.Caption = rs1!Field15
Call SealA
Call SealB
'FrmOP.Label6.Caption = rs1!Field17
FrmOP.Text12.Text = rs1!Field3
FrmOP.Text22.Text = rs1!Field4
'FrmOP.Text24.Text = rs1!Field1(41)
'FrmOP.Text25.Text = rs1!Field1(36)

'Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(6)
'CrimpStandart.Text3.Text = DataGrid1.Columns(42)
'CrimpStandart.Text1.Text = DataGrid1.Columns(10)
'CrimpStandart.Text2.Text = DataGrid1.Columns(11)
    
SetTimer hwnd, NV_CLOSEMSGBOX, 1000&, AddressOf TimerProc

  'Call MessageBox(hWnd, "Data Berhasil di Tampilkan", _
   '   MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else
'MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi"

 ' SetTimer hWnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  'Call MessageBox(hWnd, "Data yang anda cari tidak ada", _
   '   MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
   
FrmOP.Label25.Caption = ""
FrmOP.Text18.Text = ""
FrmOP.Text19.Text = ""
FrmOP.Text20.Text = ""
FrmOP.Label39.Caption = ""
FrmOP.Label40.Caption = ""
   
rs1.Close
End If

End Sub

Public Sub PANGGILWAKTU()

Form_Activate
rs2.Open "select * from LKO where AkhirPengerjaan = '" & Text1 & "' and Sequen = '" & Text3 & "' and UrutanKanban = '" & Text4 & "' ", con
If Not rs2.EOF Then
Text4.BackColor = vbBlue
 FrmOP.Command4.Caption = "Sudah Disimpan"
 FrmOP.Command4.BackColor = vbYellow
 FrmOP.Command4.Enabled = False
    
'SetTimer hwnd, NV_CLOSEMSGBOX, 1000&, AddressOf TimerProc

'Call MessageBox(hwnd, "Data sudah Tersimpan", _
'     MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else
'MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi"

 Text4.BackColor = vbRed
 FrmOP.Command4.Caption = "Simpan"
 FrmOP.Command4.Enabled = True
 FrmOP.Command4.BackColor = vbRed
 
  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

 ' Call MessageBox(hWnd, "Data Belum Tersimpan", _
  '    MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
'rs2.Close
End If

End Sub

Private Sub Text2_Change()

'Dim hitung As Integer
'hitung = Len(Text2.Text)
'If hitung = 4 Then
'Call PANGGILKANBAN2
'End If

End Sub

Private Sub SealA()
If FrmOP.Label39.Caption = "Y" Then
FrmOP.Label5.Caption = ""
Else
FrmOP.Label5.Caption = rs1!Field16
End If
End Sub

Private Sub SealB()
If FrmOP.Label40.Caption = "Y" Then
FrmOP.Label6.Caption = ""
Else
FrmOP.Label6.Caption = rs1!Field17
End If
End Sub

Private Sub Text4_Change()
'Dim hitung As Integer
'hitung = Len(Text3.Text)
'Text7 = hitung
'If hitung = 2 Then
Call PANGGILWAKTU
Call PANGGILKANBAN2

'End If
End Sub
