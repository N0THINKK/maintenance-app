VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form DataHistryOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record History Mesin"
   ClientHeight    =   6870
   ClientLeft      =   210
   ClientTop       =   555
   ClientWidth     =   9870
   ControlBox      =   0   'False
   Icon            =   "DataHistryOP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   693.94
   ScaleMode       =   0  'User
   ScaleWidth      =   100.715
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
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
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   90
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
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
      Left            =   8400
      TabIndex        =   5
      Top             =   240
      Width           =   1215
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
      Left            =   1320
      TabIndex        =   2
      Top             =   720
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
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1560
      Width           =   855
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "DataHistryOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Dim strkon As String
Dim SQL As String

Private Sub Command1_Click()
Form_Activate
rs1.CursorLocation = adUseClient
rs1.Open "select * from HS where TANGGAL = '" & Text1 & "'", kon
If Not rs1.EOF Then
    Text2 = rs1!NoMesin
         With DataGrid1
   Set .DataSource = rs1
       .Refresh
       End With
    
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Ada Informasi!!!!!"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

 

Private Sub Command3_Click()
Dim hps As String
hps = "delete from HS where TANGGAL = '" & Text1 & "'"
kon.Execute hps
MsgBox "Data Berhasil di Hapus", vbInformation, "we ada informasi"
Text1 = ""
Text2 = ""
Form_Activate
End Sub

Private Sub Command4_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from HS", kon, adOpenDynamic, adLockOptimistic
If Text1 = DataGrid1.Columns(0) Then
'Combo1.Text = DataGrid1.Columns(14)
   MsgBox "Data sudah diperiksa Leader", vbInformation, "Informasi"
Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
rs1.Update
Unload Me
DataHistry.Show


End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
Text2 = DataGrid1.Columns(1)
End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Histry.mdb"

If Text3 = "4M" Then
FrmUtama.CommandButton13.Visible = True
Else
FrmUtama.CommandButton13.Visible = False
End If

End Sub

Private Sub Form_Load()
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\Histry.mdb"
Call Histry

DataGrid1.Refresh
DoEvents

rs1.Sort = DataGrid1.Columns(2).DataField & " DESC"

Text3 = DataGrid1.Columns(14)
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Visible = False

End Sub

Private Sub Histry()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from HS", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
End Sub
