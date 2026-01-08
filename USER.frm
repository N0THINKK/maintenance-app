VERSION 5.00
Begin VB.Form USER 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5295
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   7260
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "USER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Strip Depth"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      Picture         =   "USER.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Input ID PIC"
      DownPicture     =   "USER.frx":117D3
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Picture         =   "USER.frx":12AD2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Input Data History"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5040
      Picture         =   "USER.frx":13DD1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2055
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
      Left            =   5880
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nomer Mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Ganti Password"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Picture         =   "USER.frx":15129
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "jam"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "MAINTENANCE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Paswd.Show
End Sub

Private Sub Command2_Click()
NoMesin.Show
End Sub

Private Sub Command3_Click()
FrmUtama.Show
Unload Me

End Sub

Private Sub Command4_Click()
ItemHistry.Show
End Sub

Private Sub Command5_Click()
ItemPIC.Show
End Sub

Private Sub Command6_Click()
'StripDepth.Show
End Sub

Private Sub Form_Activate()
Module1.HideXCloseButton Me

Label2 = Format(Now, "yyyy/mm/dd hh : mm : ss")

End Sub

