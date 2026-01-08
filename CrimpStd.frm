VERSION 5.00
Begin VB.Form CrimpStandart 
   AutoRedraw      =   -1  'True
   Caption         =   "Crimping Standart"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18105
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
   ScaleHeight     =   7590
   ScaleWidth      =   18105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "1 Sisi"
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
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      DataSource      =   "file1"
      Height          =   6180
      Left            =   9000
      ScaleHeight     =   6120
      ScaleWidth      =   8940
      TabIndex        =   13
      Top             =   1320
      Width           =   9000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      DataSource      =   "file1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   0
      ScaleHeight     =   6120
      ScaleWidth      =   8940
      TabIndex        =   3
      Top             =   1320
      Width           =   9000
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
      TabIndex        =   19
      Top             =   120
      Width           =   1215
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
      Height          =   405
      Left            =   14760
      TabIndex        =   17
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   15
      Top             =   840
      Width           =   1455
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
      Left            =   16680
      TabIndex        =   14
      Top             =   120
      Width           =   1215
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
      Height          =   405
      Left            =   8040
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   16560
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
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
      Left            =   14760
      TabIndex        =   4
      Text            =   ".jpg"
      Top             =   120
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
      Left            =   3360
      TabIndex        =   1
      Top             =   4680
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
      Left            =   3600
      TabIndex        =   0
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   14760
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   "CRIMPING STANDART"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
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
      Left            =   8160
      TabIndex        =   11
      Top             =   600
      Width           =   1455
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
      Height          =   255
      Left            =   16560
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
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
      TabIndex        =   5
      Top             =   5400
      Width           =   6975
   End
End
Attribute VB_Name = "CrimpStandart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim file As String

Dim myPic As PictureBox
Dim stdPic As New StdPicture

Private Sub Command1_Click()
FrmOP.Check2 = 0
FrmUtama.CheckBox1 = False

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me
CrimpStandart.Show

CrimpStandart.Text3 = FrmOP.Text33
CrimpStandart.Text4 = FrmOP.Text31
CrimpStandart.Text5 = FrmOP.Text32
CrimpStandart.Text1 = FrmOP.Text29
CrimpStandart.Text2 = FrmOP.Text30
'CrimpStandartFull.Text3 = FrmOP.Text33
'CrimpStandartFull.Text4 = FrmOP.Text31
'CrimpStandartFull.Text5 = FrmOP.Text32
'CrimpStandartFull.Text1 = FrmOP.Text29
'CrimpStandartFull.Text2 = FrmOP.Text30

'Call PanggilGambar1
'Call PanggilGambar2

End Sub

Private Sub Command3_Click()
Tampil1

Unload Me
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

Private Sub Form_Activate()
CrimpStandart.Text3 = FrmOP.Text33
CrimpStandart.Text4 = FrmOP.Text31
CrimpStandart.Text5 = FrmOP.Text32
CrimpStandart.Text1 = FrmOP.Text29
CrimpStandart.Text2 = FrmOP.Text30
FrmOP.Label46 = 0


End Sub

Private Sub Form_Load()
Combo1.AddItem "png"
Combo1.AddItem "gif"
Combo1.AddItem "bmp"
Combo1.AddItem "svg"
Combo1.AddItem "jpg"

'Text3.Text = FrmOP.Text33
'Text1.Text = FrmOP.Text29
'Text2.Text = FrmOP.Text30
'Text4.Text = FrmOP.Text31
'Text5.Text = FrmOP.Text32

End Sub


Private Sub Combo1_Click()
File1.Pattern = "*.dat;*." & Combo1.Text
End Sub

Private Sub Text1_Change()
'Dim hitung As Integer
'hitung = Len(Text1.Text)
'If hitung = 10 Then
Call PanggilGambar1
'End If
End Sub

Private Sub Text2_Change()
'Dim hitung As Integer
'hitung = Len(Text2.Text)
'If hitung = 10 Then
Call PanggilGambar2
'End If
End Sub

Public Sub PanggilGambar1()

Dim oFso As New Scripting.FileSystemObject

If oFso.FileExists(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text) Then
Me.Picture1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)

Picture1.ScaleMode = 3
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight
Picture1.AutoRedraw = True
Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / NoMesin.Text18.Text, _
    Picture1.Picture.Height / NoMesin.Text17.Text
    
    Else

Me.Picture1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Strip.jpg")
    
End If
    
End Sub

Public Sub PanggilGambar2()

Dim iFso As New Scripting.FileSystemObject

If iFso.FileExists(App.Path & "\Gambar" & "\" & Text2.Text & Text5.Text & Text3.Text & Combo1.Text) Then
Me.Picture2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text2.Text & Text5.Text & Text3.Text & Combo1.Text)

Picture2.ScaleMode = 3
Picture2.Width = Me.ScaleWidth
Picture2.Height = Me.ScaleHeight
    Picture2.AutoRedraw = True
    Picture2.PaintPicture Picture2.Picture, _
    0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
    0, 0, Picture2.Picture.Width / NoMesin.Text18.Text, _
    Picture2.Picture.Height / NoMesin.Text17.Text
        
    Else
      Me.Picture2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Strip.jpg")
 
End If
End Sub

Private Sub Tampil1()

If NoMesin.Text14 <> "" And NoMesin.Text15 <> "" Then
CrimpStandartFull.Show
End If
End Sub

