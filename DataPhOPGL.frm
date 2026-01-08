VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form DataPhOPGL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Patroli Harian Operator"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   ControlBox      =   0   'False
   Icon            =   "DataPhOPGL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8916
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cari "
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
      Left            =   4560
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   8760
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "NIK GL"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. MESIN"
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "DataPhOPGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oConn As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim strConn As String
Dim SQL As String

Sub Open_Connection()
Set oConn = New ADODB.Connection
oConn.ConnectionString = strConn
oConn.Open
End Sub

Private Sub Command1_Click()
Dim hps As String
hps = "delete from tb1 where TANGGAL = '" & Text1 & "'"
oConn.Execute hps
MsgBox "Data Berhasil di Hapus", vbInformation, "we ada informasi"
Text1 = ""
Text2 = ""
Form_Activate
End Sub

Private Sub Command2_Click()
Set rsData = New ADODB.Recordset
rsData.Open "select*from tb1", oConn, adOpenDynamic, adLockOptimistic
If Text1 = DataGrid1.Columns(0) Then
DataGrid1.Columns(26) = Combo1.Text
   MsgBox "Data sudah diperiksa Leader", vbInformation, "Informasi"
Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
rsData.Update

DataPhOPGL.Refresh

Unload Me

DataPhOPGL.Show

End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
'DataGrid1.ForeColor = vbWhite
'DataGrid1.BackColor = vnblue
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form_Activate
rsData.CursorLocation = adUseClient
rsData.Open "select * from tb1 where TANGGAL = '" & Text1 & "'", oConn
If Not rsData.EOF Then
    Text2 = rsData!NoMesin
         With DataGrid1
   Set .DataSource = rsData
       .Refresh
       End With
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Ada Informasi!!!!!"
End If
End Sub

Private Sub DataGrid1_DblClick()
DataGrid1.Columns(26) = Combo1
End Sub

Private Sub Form_Activate()
Set oConn = New ADODB.Connection
Set rsData = New ADODB.Recordset
oConn.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPHOP.mdb"
End Sub


Private Sub Form_Load()
Set oConn = New ADODB.Connection
oConn.CursorLocation = adUseClient
oConn.Provider = "Microsoft.jet.oledb.4.0"
oConn.Open App.Path & "\DbPHOP.mdb"
Call PhOP

DataGrid1.Refresh
DoEvents

rsData.Sort = DataGrid1.Columns(0).DataField & " DESC"

Combo1 = DataGrid1.Columns(26)

Dim warna As Boolean
If DataGrid1.Columns(24) = "" Then
warna = False
Else
warna = True
'Row!Tanggal(1) = vnGreen
End If


End Sub

Private Sub PhOP()
Set rsData = New ADODB.Recordset
rsData.Open "select*from tb1", oConn, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rsData
End Sub

 Private Sub Form_Resize()
On Error GoTo err
If Me.Width >= (1400 * Screen.TwipsPerPixelX) Then
   Me.Width = (1400 * Screen.TwipsPerPixelX)
   End If
   If Me.Height >= (700 * Screen.TwipsPerPixelX) Then
      Me.Height = (700 * Screen.TwipsPerPixelX)
      End If
      Exit Sub
err:
    Me.WindowState = 0
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
rsData.Sort = DataGrid1.Columns(0).DataField
End Sub

