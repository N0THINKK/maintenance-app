VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form StripDepth 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   255
   ClientWidth     =   8145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   615
      Left            =   2760
      TabIndex        =   21
      Top             =   7560
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
      Left            =   6960
      TabIndex        =   20
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1200
      TabIndex        =   16
      Top             =   7680
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4200
      TabIndex        =   14
      Top             =   960
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   8115
      TabIndex        =   13
      Top             =   1320
      Width           =   8175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
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
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
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
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Text            =   ".jpg"
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   360
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      Caption         =   "Lebar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000B&
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "_"
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipe Mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "NILAI KEDALAMAN STRIPPING"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kombinasi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No Mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "StripDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Dim strkon As String

Public jon As New ADODB.Connection
Public rs4 As New ADODB.Recordset

Dim file As String

Dim myPic As PictureBox
Dim stdPic As New StdPicture

Private Sub Combo2_Click()
Call PanggilGambar1
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Unload Me

StripDepth.Show


End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command3_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Strip", kon, adOpenDynamic, adLockOptimistic
If Text3 <> "" Then
DataGrid1.Columns(1) = Text3.Text
DataGrid1.Columns(2) = Text2.Text

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
rs1.Update
Unload Me
StripDepth.Show
End Sub

 Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

 Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Me.Picture1.Picture = LoadPicture(File1.FileName)

file = Dir1.Path & "\" & File1.FileName

Label1.Caption = file

Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / 26.46, _
    Picture1.Picture.Height / 26.46
    
    Picture1.Picture = Picture1.Image

End Sub

Private Sub Form_Load()
Combo1.AddItem ".gif"
Combo1.AddItem ".bmp"
Combo1.AddItem ".svg"
Combo1.AddItem ".jpg"

Picture1.Visible = True

Call PanggilKabel

Call PanggilKabel

Text1.Text = NoMesin.Text21.Text

Text4.Text = Left$(Text1.Text, 4)

Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\UkuranStrip.mdb"
Call Histry
rs1.Sort = DataGrid1.Columns(0).DataField & " DESC"
Text3.Text = DataGrid1.Columns(1)
Text2.Text = DataGrid1.Columns(2)

End Sub

Private Sub Histry()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Strip", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1

End Sub


Private Sub Combo1_Click()
File1.Pattern = "*.dat;*." & Combo1.Text
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Text2_Change()
'Dim hitung As Integer
'hitung = Len(Text2.Text)
'If hitung >= 8 Then
'Call PanggilGambar2
'End If
End Sub

Public Sub PanggilGambar1()

Dim oFso As New Scripting.FileSystemObject

If oFso.FileExists(App.Path & "\StripDepth" & "\" & Text4.Text & Label3 & Combo2.Text & Combo1.Text) Then
Me.Picture1.Picture = LoadPicture(App.Path & "\StripDepth" & "\" & Text4.Text & Label3 & Combo2.Text & Combo1.Text)
'Me.Image1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)

Picture1.ScaleMode = 3
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight
Picture1.AutoRedraw = True
Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / StripDepth.Text3.Text, _
    Picture1.Picture.Height / StripDepth.Text2.Text
    
    Else

Me.Picture1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Tidak.jpg")
 'Me.Image1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)
   
End If
    
End Sub

Private Sub Tampil2()

If NoMesin.Text19 <> "" And NoMesin.Text20 <> "" Then
CrimpStandart.Show
End If
End Sub

Public Sub PanggilKabel()

Set jon = New ADODB.Connection
Set rs4 = New ADODB.Recordset
jon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs4.Open "select*from Kabel", jon, adOpenDynamic, adLockOptimistic
Combo2.Clear
Do While Not rs4.EOF
Combo2.AddItem rs4!Wire
rs4.MoveNext

Loop
End Sub

Private Sub Picture1_Click()
Call PanggilGambar1
End Sub
