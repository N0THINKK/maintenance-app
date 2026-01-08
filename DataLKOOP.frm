VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form DataLKOOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lembar Kerja Operator"
   ClientHeight    =   6870
   ClientLeft      =   210
   ClientTop       =   555
   ClientWidth     =   9870
   ControlBox      =   0   'False
   Icon            =   "DataLKOOP.frx":0000
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
      TabIndex        =   14
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
   Begin VB.TextBox Text7 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   2640
      Width           =   1575
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
      TabIndex        =   4
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
   Begin VB.TextBox Text6 
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
      Left            =   4440
      TabIndex        =   11
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "QTY"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
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
      TabIndex        =   5
      Top             =   720
      Width           =   705
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
Attribute VB_Name = "DataLKOOP"
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
rs1.Open "select * from LKO where Tanggal = '" & Text1 & "' and Sequen = '" & Text2 & "'", kon
If Not rs1.EOF Then
        With DataGrid1
   Set .DataSource = rs1
       .Refresh
       End With
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi!!!!!"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

 

Private Sub Command3_Click()
Dim hps As String
hps = "delete from LKO where Tanggal = '" & Text1 & "'"
kon.Execute hps
MsgBox "Data Berhasil di Hapus", vbInformation, "informasi"
Text1 = ""
Text2 = ""
Form_Activate
End Sub

Private Sub Command4_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from LKO", kon, adOpenDynamic, adLockOptimistic
If Text1 = DataGrid1.Columns(0) Then
'DataGrid1.Columns(22) = Combo1.Text
   MsgBox "Data sudah diperiksa Leader", vbInformation, "Informasi"
Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
rs1.Update
Unload Me
DataLKO.Show


End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(2)
Text4 = DataGrid1.Columns(4)
Text3 = DataGrid1.Columns(3)
Text2 = DataGrid1.Columns(6)
Text5 = DataGrid1.Columns(7)
Text6 = DataGrid1.Columns(8)
Text7 = DataGrid1.Columns(34)
Form_Activate
Dim total
Dim total1

rs1.Open "select * from LKO where Tanggal = '" & Text1 & "' and NoMesin = '" & Text3 & "' and Shift = '" & Text4 & "'", kon
rs1.MoveFirst
total = 0
Do While Not rs1.EOF
    total = total + rs1!QtyProduct
   rs1.MoveNext
Loop
Label3 = total
rs1.Close

rs1.Open "select * from LKO where Tanggal = '" & Text1 & "' and NoMesin = '" & Text3 & "' and Shift = '" & Text4 & "' and NoLogin = '" & Text7 & "'", kon
rs1.MoveFirst
total1 = 0
Do While Not rs1.EOF
    total1 = total1 + rs1!QtyProduct
   rs1.MoveNext
Loop
Label5 = total1
rs1.Close

End Sub
Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

End Sub

Private Sub Form_Load()
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\DbLKO.mdb"
Call DbLKO

DataGrid1.Refresh
DoEvents

rs1.Sort = DataGrid1.Columns(36).DataField & " DESC"
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Visible = False

End Sub

Private Sub DbLKO()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from LKO", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
End Sub

