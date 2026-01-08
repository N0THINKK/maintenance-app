VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form NoMesin 
   Caption         =   "Pengaturan Layar"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5970
   ControlBox      =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5970
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   120
      TabIndex        =   41
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
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
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   31
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   29
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   23
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
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
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   1335
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
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3600
      ScaleHeight     =   270
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   28
      Top             =   3840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   5880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   5880
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   5880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label15 
      Caption         =   "2 Sisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4560
      TabIndex        =   38
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "1 Sisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   37
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Lebar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Ukuran Crimping Standart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Lebar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   24
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Kiri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Atas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Posisi Layar LKO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Tipe Mesin"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "No Mesin"
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
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Lebar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Posisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Ukuran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   40
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "NoMesin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Dim strkon As String
Dim SQL As String

Private Sub Command4_Click()
'Form_Activate
rs1.Open "select * from Nmr where TANGGAL = '" & Text2 & "'", kon
If Not rs1.EOF Then
    Text2 = rs1!NmrMesin
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Ada Informasi!!!!!"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command1_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Nmr", kon, adOpenDynamic, adLockOptimistic
If Text1 <> "" Then
DataGrid1.Columns(0) = Text2.Text
DataGrid1.Columns(1) = Text1.Text
DataGrid1.Columns(2) = Text4.Text
DataGrid1.Columns(3) = Text7.Text
DataGrid1.Columns(4) = Text9.Text
DataGrid1.Columns(5) = Text12.Text
DataGrid1.Columns(6) = Text13.Text
DataGrid1.Columns(7) = Text14.Text
DataGrid1.Columns(8) = Text15.Text
DataGrid1.Columns(9) = Text19.Text
DataGrid1.Columns(10) = Text20.Text

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
rs1.Update
Unload Me
NoMesin.Show

End Sub

Private Sub DataGrid1_Click()
'Text2 = DataGrid1.Columns(0)
Text1 = DataGrid1.Columns(1)
Text4 = DataGrid1.Columns(2)
Text7 = DataGrid1.Columns(3)
Text9 = DataGrid1.Columns(4)
Text12 = DataGrid1.Columns(5)
Text13 = DataGrid1.Columns(6)
Text14 = DataGrid1.Columns(7)
Text15 = DataGrid1.Columns(8)
Text19 = DataGrid1.Columns(9)
Text20 = DataGrid1.Columns(10)
End Sub

Private Sub Form_Activate()
Text2 = Format(Now, "yyyy/mm/dd hh:mm:ss")
End Sub

Private Sub Form_Load()
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\NmrMesin.mdb"
Call Histry

rs1.Sort = DataGrid1.Columns(0).DataField & " DESC"
Text3.Text = DataGrid1.Columns(7)
Text5.Text = DataGrid1.Columns(8)
Text6.Text = DataGrid1.Columns(3)
Text8.Text = DataGrid1.Columns(4)
Text10.Text = DataGrid1.Columns(5)
Text11.Text = DataGrid1.Columns(6)
Text16.Text = DataGrid1.Columns(2)
Text17.Text = DataGrid1.Columns(10)
Text18.Text = DataGrid1.Columns(9)
Text21 = DataGrid1.Columns(1)

'Text2 = DataGrid1.Columns(0)
Text1 = DataGrid1.Columns(1)
Text4 = DataGrid1.Columns(2)
Text7 = DataGrid1.Columns(3)
Text9 = DataGrid1.Columns(4)
Text12 = DataGrid1.Columns(5)
Text13 = DataGrid1.Columns(6)
Text14 = DataGrid1.Columns(7)
Text15 = DataGrid1.Columns(8)
Text19 = DataGrid1.Columns(9)
Text20 = DataGrid1.Columns(10)

End Sub

Private Sub Histry()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Nmr", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
End Sub

