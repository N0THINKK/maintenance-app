VERSION 5.00
Begin VB.Form CrimpStandartFull 
   AutoRedraw      =   -1  'True
   Caption         =   "Crimping Standart"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   600
   ClientWidth     =   13080
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
   ScaleHeight     =   9210
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   12945
      TabIndex        =   22
      Top             =   1440
      Width           =   13005
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   12915
      TabIndex        =   21
      Top             =   1320
      Width           =   12975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2 Sisi"
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
      Left            =   1680
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sisi B"
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sisi A"
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   720
      Width           =   1215
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
      TabIndex        =   17
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
      Left            =   9720
      TabIndex        =   15
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
      Left            =   1680
      TabIndex        =   13
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
      Left            =   11520
      TabIndex        =   12
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
      Left            =   5640
      TabIndex        =   7
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
      Left            =   11280
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   840
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
      Left            =   9720
      TabIndex        =   3
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
      Left            =   9720
      TabIndex        =   16
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
      TabIndex        =   14
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
      Left            =   4560
      TabIndex        =   11
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
      Left            =   5880
      TabIndex        =   10
      Top             =   600
      Width           =   1215
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
      Left            =   11280
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "CrimpStandartFull"
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
CrimpStandartFull.Show

'CrimpStandart.Text3 = Text33
'CrimpStandart.Text4 = Text31
'CrimpStandart.Text5 = Text32
'CrimpStandart.Text1 = Text29
'CrimpStandart.Text2 = Text30
CrimpStandartFull.Text3 = FrmOP.Text33
CrimpStandartFull.Text4 = FrmOP.Text31
CrimpStandartFull.Text5 = FrmOP.Text32
CrimpStandartFull.Text1 = FrmOP.Text29
CrimpStandartFull.Text2 = FrmOP.Text30

'Call PanggilGambar1
'Call PanggilGambar2

End Sub

Private Sub Command3_Click()
Picture2.Visible = False
Picture1.Visible = True
'Image2.Visible = False
'Image1.Visible = True
'Label3.Visible = False
'Label7.Visible = False
'Text5.Visible = False
'Text2.Visible = False

End Sub

Private Sub Command4_Click()
Picture2.Visible = True
Picture1.Visible = False
'Image2.Visible = True
'Image1.Visible = False
'Label6.Visible = False
'Label2.Visible = False
'Text4.Visible = False
'Text1.Visible = False

End Sub

Private Sub Command5_Click()
Tampil2

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
CrimpStandartFull.Text3 = FrmOP.Text33
CrimpStandartFull.Text4 = FrmOP.Text31
CrimpStandartFull.Text5 = FrmOP.Text32
CrimpStandartFull.Text1 = FrmOP.Text29
CrimpStandartFull.Text2 = FrmOP.Text30

FrmOP.Label46 = 0

End Sub

Private Sub Form_Load()
Combo1.AddItem ".gif"
Combo1.AddItem ".bmp"
Combo1.AddItem ".svg"
Combo1.AddItem ".jpg"

Picture1.Visible = True
Picture2.Visible = False
Call PanggilGambar1

End Sub


Private Sub Combo1_Click()
File1.Pattern = "*.dat;*." & Combo1.Text
End Sub

Private Sub Text1_Change()
'Dim hitung As Integer
'hitung = Len(Text1.Text)
'If hitung >= 8 Then
Call PanggilGambar1
'End If
End Sub

Private Sub Text2_Change()
'Dim hitung As Integer
'hitung = Len(Text2.Text)
'If hitung >= 8 Then
Call PanggilGambar2
'End If
End Sub

Public Sub PanggilGambar1()

Dim oFso As New Scripting.FileSystemObject

If oFso.FileExists(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text) Then
Me.Picture1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)
'Me.Image1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)

Picture1.ScaleMode = 3
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight
Picture1.AutoRedraw = True
Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / NoMesin.Text3.Text, _
    Picture1.Picture.Height / NoMesin.Text5.Text
    
    Else

Me.Picture1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Strip.jpg")
 'Me.Image1.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text1.Text & Text4.Text & Text3.Text & Combo1.Text)
   
End If
    
End Sub

Public Sub PanggilGambar2()

Dim iFso As New Scripting.FileSystemObject

If iFso.FileExists(App.Path & "\Gambar" & "\" & Text2.Text & Text5.Text & Text3.Text & Combo1.Text) Then
Me.Picture2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text2.Text & Text5.Text & Text3.Text & Combo1.Text)
'Me.Image2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & Text2.Text & Text5.Text & Text3.Text & Combo1.Text)

Picture2.ScaleMode = 3
Picture2.Width = Me.ScaleWidth
Picture2.Height = Me.ScaleHeight
   Picture2.AutoRedraw = True
   Picture2.PaintPicture Picture2.Picture, _
    0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
    0, 0, Picture2.Picture.Width / NoMesin.Text3.Text, _
    Picture2.Picture.Height / NoMesin.Text5.Text
        
    Else
      Me.Picture2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Strip.jpg")
 'Me.Image2.Picture = LoadPicture(App.Path & "\Gambar" & "\" & "Strip.jpg")
 
End If
End Sub

Private Sub Tampil2()

If NoMesin.Text19 <> "" And NoMesin.Text20 <> "" Then
CrimpStandart.Show
End If
End Sub

