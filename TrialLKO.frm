VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form TrialLKO 
   Caption         =   "Lembar Kerja Operator AC90"
   ClientHeight    =   7695
   ClientLeft      =   8115
   ClientTop       =   1470
   ClientWidth     =   10965
   ControlBox      =   0   'False
   Icon            =   "TrialLKO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   8520
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8916
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
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
   Begin VB.CheckBox Check1 
      Caption         =   "MD12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6600
      TabIndex        =   56
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   54
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   52
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "TrialLKO.frx":1084A
      Left            =   5160
      List            =   "TrialLKO.frx":1084C
      TabIndex        =   50
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H80000004&
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
      Height          =   330
      Left            =   8880
      TabIndex        =   49
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H80000004&
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
      Height          =   330
      Left            =   600
      TabIndex        =   48
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   46
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   45
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   43
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   42
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "TrialLKO.frx":1084E
      Left            =   4440
      List            =   "TrialLKO.frx":1085B
      TabIndex        =   36
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   35
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text13 
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
      Left            =   240
      TabIndex        =   34
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "TrialLKO.frx":10869
      Left            =   2280
      List            =   "TrialLKO.frx":10888
      TabIndex        =   32
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   29
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   28
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   26
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text8 
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
      Left            =   6480
      TabIndex        =   24
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text7 
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
      Left            =   3120
      TabIndex        =   23
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text6 
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
      Left            =   3120
      TabIndex        =   22
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Left            =   3120
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
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
      Left            =   3120
      TabIndex        =   20
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
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
      Left            =   3120
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SIMPAN"
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
      Left            =   360
      TabIndex        =   6
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9240
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   58
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label29 
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   59
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label27 
      Caption         =   "No. 4M"
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
      Left            =   240
      TabIndex        =   55
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "C/L"
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
      Left            =   3840
      TabIndex        =   53
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Defect"
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
      Left            =   5160
      TabIndex        =   51
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      Picture         =   "TrialLKO.frx":108DF
      Top             =   1680
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "TrialLKO.frx":10D5F
      Top             =   1680
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      X1              =   1800
      X2              =   9000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label26 
      Caption         =   "Lot ID Wire"
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
      Left            =   5520
      TabIndex        =   47
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Front C/H"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   41
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "Front C/W"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   40
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Rear C/H"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   39
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Rear C/W"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   38
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   37
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "No Mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   33
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "NIK"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   31
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "QTY PRODUK"
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
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "QTY DEFECT"
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
      Left            =   2160
      TabIndex        =   25
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Rear C/W"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Rear C/H"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Front C/W"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Front C/H"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Lot ID Term B"
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
      Left            =   5520
      TabIndex        =   14
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Lot ID Term A"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "SISI B"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "SISI A"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "LEMBAR KERJA OPERATOR"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "SealB"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "SealA"
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
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEQUEN"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   825
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   1560
      X2              =   9240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label25 
      Caption         =   "Kombinasi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   44
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "TrialLKO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Public Con As New ADODB.Connection
Public rs2 As New ADODB.Recordset

Public Bon As New ADODB.Connection
Public rs3 As New ADODB.Recordset

Public son As New ADODB.Connection
Public rs4 As New ADODB.Recordset

Private Sub Check1_Click()
'Call MD12
If Text5 = "" Then
MSComm1.Output = Text5.Text
Else
If Text6 = "" Then
MSComm1.Output = Text6.Text
Else
If Text7 = "" Then
MSComm1.Output = Text7.Text
Else
If Text8 = "" Then
MSComm1.Output = Text8.Text
Else
If Text9 = "" Then
MSComm1.Output = Text9.Text
Else
If Text10 = "" Then
MSComm1.Output = Text10.Text
Else
If Text14 = "" Then
MSComm1.Output = Text14.Text
Else
If Text15 = "" Then
MSComm1.Output = Text15.Text
End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Command1_Click()

'Form_Activate
'rs1.Open "select * from Product where Sequen = '" & Text13 & "'", kon
'If Not rs1.EOF Then
  '  Text1 = rs1!Sequen
  '  Text17 = rs1!SemiStripA
  '  Label25 = rs1!Kombinasi
   ' Text18 = rs1!TermA
  '  Text19 = rs1!TermB
   ' Text20 = rs1!Field5
   ' Label5 = rs1!NmrSealA
  '  Label6 = rs1!NmrSealB
    
'MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
'Else
  '  MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi!!!!!"
    'rs1.Close
'End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim hps As String
hps = "delete from Product where Sequen = '" & Text1 & "'"
kon.Execute hps
MsgBox "Data Berhasil di Hapus", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Form_Activate
End Sub

Private Sub Command4_Click()

If Text11 = "" Then
MsgBox " Data Tanggal belum terisi"
Exit Sub
End If

If Combo1 = "" Then
MsgBox " No Mesin Belum terisi"
Exit Sub
End If

If Combo2 = "" Then
MsgBox " Shift belum terisi"
Exit Sub
End If

If Combo2 = "" Then
MsgBox " NIK belum terisi"
Exit Sub
End If

If Text1 = "" Then
MsgBox " Sequen belum dipilih"
Exit Sub
End If

If Text17 = "" Then
MsgBox " QTY Produk belum terisi"
Exit Sub
End If

Dim IsiA As Boolean
If Text3 = "" Then
IsiA = True
Else
IsiA = False
End If

If IsiA = True Then
Text4 = ""
Else
MsgBox " Terminal sisi A belum terisi"
Exit Sub
End If


Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into LKO (Tanggal,NoMesin,Shift,NIK,Sequen,CutL,LotIDWire,Kombinasi,TermA,SealA,TermB,SealB,LotIDTermA,FCHTermA,FCWTermA,RCHTermA,RCWTermA,LotIDTermB,FCHTermB,FCWTermB,RCHTermB,RCWTermB,KodeDefect,QtyDefect,QtyProduct,No4M) values ('" _
& Text11 & "','" & Combo1 & "','" & Combo2 & "','" & Combo2 & "','" & Text1 & "','" & Text20 & "','" & Text2 & "','" & Label25 & "','" & Label3 & "','" & Label5 & "','" & Label4 & "','" & Label6 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "','" & Text8 & "','" & Text9 & "','" & Text10 & "','" & Text14 & "','" & Text15 & "','" & Combo3 & "','" & Text16 & "','" & Text17 & "','" & Text21 & "')"

Con.Execute SQLTambah

MsgBox "Input data berhasil", vbDefaultButton1

'Combo1 = ""
'Combo2 = ""
'Combo2 = ""
Text1 = ""
Text2 = ""
Label25 = ""
Label3 = ""
Label5 = ""
Label4 = ""
Label6 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text14 = ""
Text15 = ""
Combo3 = ""
Text16 = ""
Text17 = ""

'rs2.Update

Form_Activate
rs2.Open "select * from LKO where Tanggal = '" & Text11 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'", Con
rs2.MoveFirst
Total = 0
Do While Not rs2.EOF
    Total = Total + rs2!QtyProduct
   rs2.MoveNext
Loop
Label28 = Total

End Sub

Private Sub Command5_Click()
DataLKOOP.Show
End Sub

Private Sub Command6_Click()
Form_Activate
rs2.Open "select * from LKO where Sequen = '" & Text13 & "'", Con
If Not rs2.EOF Then
   Text2 = rs2!NoMesin
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Informasi!!!!!"
    rs2.Close
End If
End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(6)
'Label25.Caption = DataGrid1.Columns(42)
'Text18.Text = DataGrid1.Columns(10)
'Text19.Text = DataGrid1.Columns(11)
'Text20.Text = DataGrid1.Columns(4)
'Label5.Caption = DataGrid1.Columns(14)
'Label6.Caption = DataGrid1.Columns(15)

'If Text16 = "" Then
'Text16.Text = 0
Exit Sub
End If

End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Produk.mdb"

Set Con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
Con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

Set Bon = New ADODB.Connection
Set rs3 = New ADODB.Recordset
Bon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPrdMaster.mdb"

Set son = New ADODB.Connection
Set rs4 = New ADODB.Recordset
Bon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbCHCW.mdb"

Combo1.Text = NoMesin.Text3.Text

Dim henti As Boolean
henti = False
'Do Until henti
Text11 = Format(Now, "yyyy/mm/dd")
DoEvents
'Loop

If "2" = DataGrid1.Columns(14) Then
Image2.Visible = True
Else
Image2.Visible = False
End If

End Sub


Private Sub Form_Load()
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\Produk.mdb"
Call LKO
'rs1.Sort = DataGrid1.Columns(0).DataField

Set Con = New ADODB.Connection
Con.CursorLocation = adUseClient
Con.Provider = "Microsoft.jet.oledb.4.0"
Con.Open App.Path & "\DbLKO.mdb"
Call SimpanLKO
'rs2.Sort = DataGrid2.Columns(0).DataField & " DESC"


   ' With MSComm1
   '     .CommPort = 3
    '    .Settings = "9600,n,8,1"
    '    .Handshaking = comNone
    '    .SThreshold = 0 'No events after send completions.
    '    .RThreshold = 0 'No events after receive completions.
    '    .PortOpen = True
   ' End With

With Combo3
.AddItem "A.1 Core Terurai"
.AddItem "A.2 Core Terpotong"
.AddItem "A.3 Core Rusak"
.AddItem "A.4 Core Tidak Teratur"
.AddItem "A.5 Core Maju"
.AddItem "A.6 Core Terurai"
.AddItem "A.7 Core Terpotong"
.AddItem "A.8 Core Rusak"
.AddItem "B.1 Terminal Tergores"
.AddItem "B.2 Terminal Bengkok ke atas"
.AddItem "B.3 Terminal Bengkok ke bawah"
.AddItem "B.4 Terminal Melintir"
.AddItem "B.5 Terminal Ujung Terpotong"
.AddItem "B.6 Terminal Ujung Terbuka"
.AddItem "B.7 Terminal Ujung Rusak"
.AddItem "B.8 Terminal Bridge terlalu panjang"
.AddItem "B.9 Terminal Rusak"
.AddItem "B.10 Terminal Lepas dari Circuit"
.AddItem "C.1 Front C/H terlalu tinggi"
.AddItem "C.2 Front C/H terlalu rendah"
.AddItem "C.3 Front C/W terlalu tinggi"
.AddItem "C.4 Front C/W terlalu rendah"
.AddItem "C.5 Front Flash"
.AddItem "D.1 Rear C/H terlalu tinggi"
.AddItem "D.2 Rear C/H terlalu rendah"
.AddItem "D.3 Rear C/W terlalu tinggi"
.AddItem "D.4 Rear C/W terlalu rendah"
.AddItem "D.5 Rear ada di dalam Insulasi"
.AddItem "D.6 Rear Tidak Tercrimping"
.AddItem "D.7 Rear Tidak seimbang"
.AddItem "E.1 Insulation Tercrimping"
.AddItem "E.2 Insulation Terlalu mundur"
.AddItem "E.3 Insulation Rusak"
.AddItem "E.4 Insulation Tidak rata"
.AddItem "F.1 Seal Terpotong"
.AddItem "F.2 Seal Terbalik"
.AddItem "F.3 Seal Terlalu mundur"
.AddItem "F.4 Seal Terlalu maju"
.AddItem "F.5 Seal Tercrimping"
.AddItem "F.6 Seal Tidak ada"
.AddItem "F.7 Seal Sobek"
.AddItem "G.1 Crimping Ada Benda Asing"
.AddItem "G.2 Crimping Ada 2 Terminal Tercrimping"
.AddItem "G.3 Crimping Tanpa Core"
.AddItem "G.4 Crimping Tanpa Stripping"
.AddItem "H.1 Lance Rusak"
.AddItem "H.2 Stabilizer Rusak"
.AddItem "H.3 Bellmouth Tidak Standart"
End With

NIKOP

End Sub

Private Sub MD12()
MSComm1.CommPort = 3 'tergantung COM port yang digunakan
MSComm1.Settings = "9600,N,8,1" 'contoh setting serial port
MSComm1.InputLen = 0
MSComm1.SThreshold = 0
MSComm1.RThreshold = 0
MSComm1.PortOpen = True
End Sub

Private Sub LKO()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Product", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
End Sub
Private Sub SimpanLKO()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
'Set DataGrid2.DataSource = rs2
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

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub MSComm1_OnComm()
If Len(Text6.Text) > 0 Then
        MSComm1.Output = Text6.Text
        Text6.Text = vbNullString
    End If
    Text6.SetFocus

End Sub

Private Sub Text18_Change()
If Text18.Text = "STRIP ONLY" Or Text18.Text = "BONDER" Then
Image1.Visible = False
Else
Image1.Visible = True
End If
End Sub

Private Sub Text19_Change()
If Text19.Text = "STRIP ONLY" Or Text19.Text = "BONDER" Then
Image2.Visible = False
Else
Image2.Visible = True
End If
End Sub

Public Sub NIKOP()

Set Don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
Don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from NIKOP", Don, adOpenDynamic, adLockOptimistic
Combo4.Clear
Do While Not rs3.EOF
Combo4.AddItem rs3!OPERATOR
rs3.MoveNext

Loop
End Sub


