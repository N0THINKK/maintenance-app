VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form ItemHistry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Item History Mesin"
   ClientHeight    =   6870
   ClientLeft      =   210
   ClientTop       =   555
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "ItemHistry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   693.94
   ScaleMode       =   0  'User
   ScaleWidth      =   121.226
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
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
      Left            =   9360
      TabIndex        =   26
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   25
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   24
      Top             =   6240
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   5175
      Left            =   6720
      TabIndex        =   22
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   5175
      Left            =   4320
      TabIndex        =   21
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   2040
      TabIndex        =   20
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
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
      TabIndex        =   19
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton Command10 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text4 
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
      Left            =   6720
      TabIndex        =   9
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text3 
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
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   6240
      Width           =   855
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
      Left            =   6480
      TabIndex        =   5
      Top             =   120
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
      Height          =   375
      Left            =   11040
      TabIndex        =   4
      Top             =   0
      Width           =   735
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
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2055
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Height          =   5175
      Left            =   9360
      TabIndex        =   23
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Wire"
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
      Left            =   9360
      TabIndex        =   27
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penyebab"
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
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Problem"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIK"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   930
   End
End
Attribute VB_Name = "ItemHistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rs4 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Dim strkon As String
Dim SQL As String



Private Sub Command10_Click()
Set rs4 = New ADODB.Recordset
rs4.Open "select*from Cara", kon, adOpenDynamic, adLockOptimistic
DataGrid4.Columns(0) = Text4.Text
   MsgBox "Data berhasil di Edit", vbInformation, "Informasi"
   
rs4.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command11_Click()
Set rs5 = New ADODB.Recordset
rs5.Open "select*from Kabel", kon, adOpenDynamic, adLockOptimistic
DataGrid5.Columns(0) = Text5.Text
   MsgBox "Data berhasil di Edit", vbInformation, "Informasi"
   
rs5.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command12_Click()
Set rs5 = New ADODB.Recordset
rs5.Open "select*from Cara", kon, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Kabel (Wire) values ('" & Text5 & "')"

kon.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

rs5.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim hps1 As String
Dim hps2 As String
Dim hps3 As String
Dim hps4 As String
Dim hps5 As String

hps1 = "delete from NIKPATROL where NIK = '" & Text1 & "'"
hps2 = "delete from Problem where PROBLEM = '" & Text2 & "'"
hps3 = "delete from Sebab where Penyebab = '" & Text3 & "'"
hps4 = "delete from Cara where Tindakan = '" & Text4 & "'"
hps5 = "delete from Kabel where Wire = '" & Text5 & "'"


kon.Execute hps1
kon.Execute hps2
kon.Execute hps3
kon.Execute hps4
kon.Execute hps5

MsgBox "Data Berhasil di Hapus", vbInformation, "informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""

Form_Activate
End Sub

Private Sub Command1_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from NIKPATROL", kon, adOpenDynamic, adLockOptimistic
DataGrid1.Columns(0) = Text1.Text
   MsgBox "Data berhasil di Edit", vbInformation, "Informasi"
   
rs1.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show

End Sub

Private Sub Command4_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from NIKPATROL", kon, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into NIKPATROL (NIK) values ('" & Text1 & "')"

kon.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

rs1.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show

End Sub

Private Sub Command5_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from Problem", kon, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Problem (PROBLEM) values ('" & Text2 & "')"

kon.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

rs2.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command6_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from Problem", kon, adOpenDynamic, adLockOptimistic
DataGrid2.Columns(0) = Text2.Text
   MsgBox "Data berhasil di Edit", vbInformation, "Informasi"
   
rs2.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command7_Click()
Set rs3 = New ADODB.Recordset
rs3.Open "select*from Sebab", kon, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Sebab (Penyebab) values ('" & Text3 & "')"

kon.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

rs3.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command8_Click()
Set rs3 = New ADODB.Recordset
rs3.Open "select*from Sebab", kon, adOpenDynamic, adLockOptimistic
DataGrid3.Columns(0) = Text3.Text
   MsgBox "Data berhasil di Edit", vbInformation, "Informasi"
   
rs3.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub Command9_Click()
Set rs4 = New ADODB.Recordset
rs4.Open "select*from Cara", kon, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Cara (Tindakan) values ('" & Text4 & "')"

kon.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

rs4.Update

ItemHistry.Refresh

Unload Me

ItemHistry.Show
End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
End Sub

Private Sub DataGrid2_Click()
Text2 = DataGrid2.Columns(0)
End Sub
Private Sub DataGrid3_Click()
Text3 = DataGrid3.Columns(0)
End Sub

Private Sub DataGrid4_Click()
Text4 = DataGrid4.Columns(0)
End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"


End Sub

Private Sub Form_Load()
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\ListHistry.mdb"
Call Histry

rs1.Sort = DataGrid1.Columns(0).DataField
rs2.Sort = DataGrid2.Columns(0).DataField
rs3.Sort = DataGrid3.Columns(0).DataField
rs4.Sort = DataGrid4.Columns(0).DataField
rs5.Sort = DataGrid5.Columns(0).DataField


Text2.Text = LCase(Text2.Text)

End Sub

Private Sub Histry()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from NIKPATROL", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1

Set rs2 = New ADODB.Recordset
rs2.Open "select*from Problem", kon, adOpenDynamic, adLockOptimistic
Set DataGrid2.DataSource = rs2

Set rs3 = New ADODB.Recordset
rs3.Open "select*from Sebab", kon, adOpenDynamic, adLockOptimistic
Set DataGrid3.DataSource = rs3

Set rs4 = New ADODB.Recordset
rs4.Open "select*from Cara", kon, adOpenDynamic, adLockOptimistic
Set DataGrid4.DataSource = rs4

Set rs5 = New ADODB.Recordset
rs5.Open "select*from Kabel", kon, adOpenDynamic, adLockOptimistic
Set DataGrid5.DataSource = rs5


End Sub

Private Sub Text2_Change()
'Text2.Text = UCase(Text2.Text)
End Sub

Private Sub Text3_Change()
'Text3.Text = StrConv(Text3.Text, vbProperCase)
End Sub

Private Sub Text4_Change()
'Text4.Text = StrConv(Text4.Text, vbProperCase)
End Sub
