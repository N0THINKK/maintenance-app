VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form DataPhOPOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Patroli Harian Operator"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   ControlBox      =   0   'False
   Icon            =   "DataPhOPOP2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      TabIndex        =   7
      Top             =   5760
      Width           =   2895
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
Attribute VB_Name = "DataPhOPOP"
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

End Sub

Private Sub Command2_Click()
End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
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

rsData.Sort = DataGrid1.Columns(0).DataField & " DESC"

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

