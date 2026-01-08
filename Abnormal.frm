VERSION 5.00
Begin VB.Form Abnormal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abnormalitas"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   7800
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   195
      Left            =   6960
      TabIndex        =   13
      Top             =   7800
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   1305
      Left            =   8630
      Stretch         =   -1  'True
      Top             =   6200
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   1300
      Left            =   7400
      Stretch         =   -1  'True
      Top             =   6200
      Width           =   1150
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIK"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   7560
      Width           =   1155
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIK"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   7400
      TabIndex        =   11
      Top             =   7560
      Width           =   1150
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "  DAN/TUM"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9260
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "       LNA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "     TWC"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7800
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  PREPARE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   9240
      TabIndex        =   7
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECKED"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   8520
      TabIndex        =   6
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "APPROVE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   7800
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REVISION        :  N"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7700
      TabIndex        =   4
      Top             =   200
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ISSUE DATE    : 17 FEBRUARI 2012 "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7700
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTROL NO : JAI/QA/ABN/PA/002"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7700
      TabIndex        =   2
      Top             =   10
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "QA DEPT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PT. JATIM AUTOCOMP INDONESIA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   7995
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Abnormal.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10005
   End
End
Attribute VB_Name = "Abnormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
DataAbnormal.Show
End Sub

Private Sub Command2_Click()
Unload Me
Abnormal.Show
End Sub

Private Sub Form_Load()
Label9.Caption = DataAbnormal.Text3
Label10.Caption = DataAbnormal.Text2
Label11.Caption = DataAbnormal.Text1
Label12.Caption = DataAbnormal.Text4
Label13.Caption = DataAbnormal.Text5

PanggilGambar1
PanggilGambar2

End Sub

Public Sub PanggilGambar1()

Dim oFso As New Scripting.FileSystemObject

If oFso.FileExists(App.Path & "\GL" & "\" & Label12.Caption & "\" & ".jpg") Then
'Me.Picture1.Picture = LoadPicture(App.Path & "\GL" & "\" & Label12.Caption & "\" & ".jpg")
Me.Image2.Picture = LoadPicture(App.Path & "\GL" & "\" & Label12.Caption & "\" & ".jpg")


'Picture1.ScaleMode = 3
'Picture1.Width = Me.ScaleWidth
'Picture1.Height = Me.ScaleHeight
'Picture1.AutoRedraw = True
'Picture1.PaintPicture Picture1.Picture, _
 '   0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
 '   0, 0, Picture1.Picture.Width / DataAbnormal.Text6.Text, _
 '   Picture1.Picture.Height / DataAbnormal.Text7.Text
    
    Else

'Me.Picture1.Picture = LoadPicture(App.Path & "\GL" & "\" & "Refresh.jpg")
Me.Image2.Picture = LoadPicture(App.Path & "\GL" & "\" & "Refresh.jpg")
    
End If
    
End Sub

Public Sub PanggilGambar2()

Dim iFso As New Scripting.FileSystemObject

If iFso.FileExists(App.Path & "\GL" & "\" & Label13.Caption & "\" & ".jpg") Then
'Me.Picture2.Picture = LoadPicture(App.Path & "\GL" & "\" & Label13.Caption & "\" & ".jpg")
Me.Image3.Picture = LoadPicture(App.Path & "\GL" & "\" & Label12.Caption & "\" & ".jpg")


'Picture2.ScaleMode = 3
'Picture2.Width = Me.ScaleWidth
'Picture2.Height = Me.ScaleHeight
'    Picture2.AutoRedraw = True
 '   Picture2.PaintPicture Picture2.Picture, _
   ' 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
  '  0, 0, Picture2.Picture.Width / DataAbnormal.Text6.Text, _
    Picture2.Picture.Height / DataAbnormal.Text7.Text
        
    Else
      'Me.Picture2.Picture = LoadPicture(App.Path & "\GL" & "\" & "Refresh.jpg")
      Me.Image3.Picture = LoadPicture(App.Path & "\GL" & "\" & "Refresh.jpg")

End If
End Sub

