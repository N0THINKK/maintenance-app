VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmOPtest 
   Caption         =   " "
   ClientHeight    =   8745
   ClientLeft      =   255
   ClientTop       =   900
   ClientWidth     =   9060
   ControlBox      =   0   'False
   Icon            =   "FrmLKOAC90.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   583
   ScaleMode       =   0  'User
   ScaleWidth      =   604
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      TabIndex        =   91
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   88
      Top             =   840
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1695
      Left            =   0
      TabIndex        =   61
      Top             =   6840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   7223
      _Version        =   393216
      AllowArrows     =   -1  'True
      ColumnHeaders   =   0   'False
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   14
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
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
      Caption         =   "Command3"
      Height          =   195
      Left            =   120
      TabIndex        =   84
      Top             =   6300
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   4095
      Left            =   7560
      TabIndex        =   72
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   7223
      _Version        =   393216
      AllowArrows     =   -1  'True
      ColumnHeaders   =   0   'False
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   14
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.TextBox Text28 
      Height          =   435
      Left            =   6360
      TabIndex        =   81
      Text            =   "Text28"
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text27 
      Alignment       =   1  'Right Justify
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
      Left            =   3480
      TabIndex        =   77
      Text            =   "0"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Height          =   435
      Left            =   7680
      TabIndex        =   76
      Text            =   "Text26"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1695
      Left            =   4680
      TabIndex        =   62
      Top             =   6840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      TabIndex        =   70
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Crimping Standart"
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
      Left            =   3480
      TabIndex        =   69
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   3480
      Top             =   7320
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Klik Terminal B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Klik Terminal A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   0
   End
   Begin VB.CommandButton Command6 
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
      Height          =   495
      Left            =   4800
      TabIndex        =   59
      Top             =   2160
      Width           =   1095
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
      Left            =   5760
      TabIndex        =   53
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   51
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FrmLKOAC90.frx":1084A
      Left            =   1320
      List            =   "FrmLKOAC90.frx":1084C
      TabIndex        =   47
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   43
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      TabIndex        =   42
      Text            =   "0"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   40
      Text            =   "0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   39
      Text            =   "0"
      Top             =   3840
      Width           =   1095
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
      ItemData        =   "FrmLKOAC90.frx":1084E
      Left            =   4440
      List            =   "FrmLKOAC90.frx":1085B
      TabIndex        =   33
      Top             =   840
      Width           =   1215
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
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   5160
      Width           =   1095
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
      ItemData        =   "FrmLKOAC90.frx":10869
      Left            =   1680
      List            =   "FrmLKOAC90.frx":10888
      TabIndex        =   30
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   26
      Text            =   "0"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   24
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   22
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   21
      Text            =   "0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   20
      Text            =   "0"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   19
      Text            =   "0"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   18
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
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
      Height          =   855
      Left            =   4920
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
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
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   2265
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5280
      TabIndex        =   58
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   57
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Left            =   3360
      TabIndex        =   49
      Top             =   1800
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
      Height          =   495
      Left            =   3360
      TabIndex        =   63
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   64
      Top             =   7680
      Width           =   855
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
      Height          =   210
      Left            =   6720
      TabIndex        =   46
      Top             =   1320
      Width           =   1455
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
      Left            =   8520
      TabIndex        =   54
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   45
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   6360
      TabIndex        =   75
      Text            =   "Text25"
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label46 
      Caption         =   "No Urut"
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
      Left            =   8760
      TabIndex        =   92
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label45 
      Height          =   255
      Left            =   1560
      TabIndex        =   90
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label44 
      Caption         =   "No Urut"
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
      Left            =   3360
      TabIndex        =   89
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label43 
      Height          =   255
      Left            =   1200
      TabIndex        =   87
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label42 
      Height          =   255
      Left            =   1440
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label41 
      Caption         =   "NoLogin"
      Height          =   255
      Left            =   480
      TabIndex        =   85
      Top             =   6240
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "FrmLKOAC90.frx":108DF
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label40 
      Caption         =   "Label40"
      Height          =   255
      Left            =   5040
      TabIndex        =   83
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label39 
      Caption         =   "Label39"
      Height          =   255
      Left            =   960
      TabIndex        =   82
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label38 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   6720
      TabIndex        =   80
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label37 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   6360
      TabIndex        =   79
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label36 
      Caption         =   "Defect Mesin"
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
      Left            =   3600
      TabIndex        =   78
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label35 
      Caption         =   "TERSIMPAN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7920
      TabIndex        =   74
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label34 
      Caption         =   "SEQUEN"
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
      Left            =   6120
      TabIndex        =   73
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Issue"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2520
      TabIndex        =   71
      Top             =   2295
      Width           =   750
   End
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7800
      TabIndex        =   68
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   67
      Top             =   1800
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   8
      X1              =   128
      X2              =   456
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Label Label30 
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   120
      Width           =   1095
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
      Height          =   375
      Left            =   7560
      TabIndex        =   55
      Top             =   840
      Width           =   1215
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
      Left            =   7560
      TabIndex        =   56
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
      Left            =   120
      TabIndex        =   52
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "C/L"
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
      Left            =   2640
      TabIndex        =   50
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Defect"
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
      Left            =   1440
      TabIndex        =   48
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Picture         =   "FrmLKOAC90.frx":10D4B
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label26 
      Caption         =   "Lot ID Wire"
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
      Left            =   120
      TabIndex        =   44
      Top             =   2640
      Width           =   1215
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
      Left            =   3720
      TabIndex        =   38
      Top             =   3480
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
      Left            =   3720
      TabIndex        =   37
      Top             =   4200
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
      Left            =   3720
      TabIndex        =   36
      Top             =   3840
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
      Left            =   3720
      TabIndex        =   35
      Top             =   4560
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
      Left            =   4440
      TabIndex        =   34
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
      Left            =   1800
      TabIndex        =   31
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
      Left            =   5760
      TabIndex        =   29
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
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "QTY PRODUK"
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
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "QTY Defect Operator"
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
      Left            =   1320
      TabIndex        =   23
      Top             =   5640
      Width           =   1815
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
      Left            =   1440
      TabIndex        =   16
      Top             =   4560
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
      Left            =   1440
      TabIndex        =   15
      Top             =   3840
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
      Left            =   1440
      TabIndex        =   14
      Top             =   4200
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
      Left            =   1440
      TabIndex        =   13
      Top             =   3480
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
      Left            =   3720
      TabIndex        =   12
      Top             =   3000
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
      Left            =   1560
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "SISI B"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "SISI A"
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
      TabIndex        =   9
      Top             =   2640
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
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "SealB"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "SealA"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEQUEN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   690
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   104
      X2              =   480
      Y1              =   112
      Y2              =   112
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
      Left            =   3720
      TabIndex        =   41
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FrmOPtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Public rs6 As New ADODB.Recordset
Public rs7 As New ADODB.Recordset

Public Con As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset

Public jon As New ADODB.Connection
Public rs4 As New ADODB.Recordset

Public don As New ADODB.Connection
Public rs8 As New ADODB.Recordset

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

Private Sub Check2_Click()
If Check2 = 1 Then
CrimpStandartFull.Show
Else
CrimpStandartFull.Hide
End If

End Sub

Private Sub Combo5_Change()

Abjad

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Kon2.Show
End Sub

Private Sub Command4_Click()

If Text11 = "" Then
MsgBox " Data Tanggal belum terisi"
Exit Sub
End If

If Text17 = "" Then
MsgBox " Qty Produk belum terisi"
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


Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into LKO (Tahun,Bulan,Tanggal,NoMesin,Shift,NIK,Sequen,UrutanKanban,Waktu,AwalPengerjaan,AkhirPengerjaan,CutL,LotIDWire,Kombinasi,TermA,SealA,TermB,SealB,LotIDTermA,FCHTermA,FCWTermA,RCHTermA,RCWTermA,LotIDTermB,FCHTermB,FCWTermB,RCHTermB,RCWTermB,KodeDefect,QtyDefectOperator,QtyDefectMesin,QtyProduct,No4M,NoLogin,NoIssue) values ('" _
& Label37 & "','" & Label38 & "','" & Text11 & "','" & Combo1 & "','" & Combo2 & "','" & Combo4 & "','" & Text1 & "','" & Text28 & "','" & Label30 & "','" & Text25 & "','" & Text26 & "','" & Text20 & "','" & Text2 & "','" & Label25 & "','" & Text18 & "','" & Label5 & "','" & Text19 & "','" & Label6 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "','" & Text8 & "','" & Text9 & "','" & Text10 & "','" & Text14 & "','" & Text15 & "','" & Combo3 & "','" & Text16 & "','" & Text27 & "','" & Text17 & "','" & Text21 & "','" & Label41 & "','" & Text24 & "')"

Con.Execute SQLTambah

  SetTimer hWnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hWnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Text1 = ""
Text2 = ""
'Label25 = ""
Label3 = ""
'Label5 = ""
Label4 = ""
'Label6 = ""
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
Text24 = ""

'Form_Activate
rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'", Con
rs2.MoveFirst
Total = 0
Do While Not rs2.EOF
    Total = Total + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'and TermB = '" & Text19 & "'", Con
rs2.MoveFirst
TeB = 0
Do While Not rs2.EOF
    TeB = TeB + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "' and TermA = '" & Text18 & "'", Con
rs2.MoveFirst
TeA = 0
Do While Not rs2.EOF
    TeA = TeA + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

Unload Me

FrmOP.Show

Text28.Text = "0"
Text28.Text = DataGrid1.Columns(1)
Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

'Label25.Caption = DataLKO.DataGrid1.Columns(8)
'Text18.Text = DataLKO.DataGrid1.Columns(9)
'Text19.Text = DataLKO.DataGrid1.Columns(11)
'Text20.Text = DataLKO.DataGrid1.Columns(6)
'Label5.Caption = DataLKO.DataGrid1.Columns(10)
'Label6.Caption = DataLKO.DataGrid1.Columns(12)


'Call NIKRE
Label28 = Total
Label31 = TeA
Label32 = TeB

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus
End Sub

Private Sub Command5_Click()
DataLKOOP.Show
End Sub

Private Sub Command6_Click()
Call MuatUlang

End Sub

Private Sub DataGrid1_Click()
Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

End Sub

Private Sub DataGrid2_Click()
Text4.Text = DataGrid2.Columns(17)
Text5.Text = DataGrid2.Columns(18)
Text6.Text = DataGrid2.Columns(21)
Text7.Text = DataGrid2.Columns(22)
End Sub

Private Sub DataGrid3_Click()
Text9.Text = DataGrid3.Columns(25)
Text10.Text = DataGrid3.Columns(26)
Text14.Text = DataGrid3.Columns(29)
Text15.Text = DataGrid3.Columns(30)
End Sub

Private Sub Form_Activate()

Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Jissk.mdb"

Set Con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"


'Label43.Caption = Label42

Combo5.Text = Right$(Combo1.Text, 2)

Dim henti As Boolean
henti = False
'Do Until henti
Text11 = Format(Now, "yyyy/mm/dd")
Label37 = Format(Now, "yyyy")
Label38 = Format(Now, "mm")
Label41 = FrmUtama.Label6
DoEvents
'Loop

Text13.Text = DataGrid2.Columns(17)
Text23.Text = DataGrid3.Columns(25)

If Text13.Text = "    0" Then
Command1.Visible = False
Else
Command1.Visible = True
End If

If Text23.Text = "    0" Then
Command7.Visible = False
Else
Command7.Visible = True
End If

With DataGrid1
.Columns(0).Width = 36
.Columns(1).Width = 36
.Columns(2).Visible = False
.Columns(3).Visible = True
.Columns(4).Visible = False
.Columns(5).Visible = False
End With

If Command4.Caption = "Sudah Disimpan" Then
Command4.MaskColor = vbRed
Else
Command4.MaskColor = vbYellow
End If

End Sub

Private Sub Form_Load()

Me.Top = NoMesin.Text8
Me.Left = NoMesin.Text6

Combo1.Text = NoMesin.Text21.Text
'Abjad

'Form_Resize
Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\Jissk.mdb"
Call LKO

Set Con = New ADODB.Connection
Con.CursorLocation = adUseClient
Con.Provider = "Microsoft.jet.oledb.4.0"
Con.Open App.Path & "\DbLKO.mdb"
Call HasilSimpan

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

'With Combo5
'.AddItem "01"
'.AddItem "02"
'.AddItem "03"
'.AddItem "04"
'.AddItem "05"
'.AddItem "06"
'.AddItem "07"
'.AddItem "08"
'.AddItem "09"
'.AddItem "10"
'.AddItem "11"
'.AddItem "12"
'.AddItem "13"
'.AddItem "14"
'.AddItem "15"
'End With

End Sub

Private Sub LKO()

Set rs1 = New ADODB.Recordset
rs1.Open "select*from Prdlog", kon, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs1
rs1.Sort = DataGrid1.Columns(3).DataField & " DESC"

Set rs5 = New ADODB.Recordset
rs5.Open "select*from Hasil", kon, adOpenDynamic, adLockOptimistic
Set DataGrid2.DataSource = rs5
rs5.Sort = DataGrid2.Columns(36).DataField & " DESC"
With DataGrid2
.Columns(0).Visible = False
.Columns(1).Visible = False
.Columns(2).Visible = True
.Columns(3).Visible = False
.Columns(4).Visible = True
.Columns(5).Visible = False
.Columns(6).Visible = False
.Columns(7).Visible = False
.Columns(8).Visible = False
.Columns(9).Visible = False
.Columns(10).Visible = False
.Columns(11).Visible = True
.Columns(12).Visible = False
.Columns(13).Visible = False
.Columns(14).Visible = False
.Columns(15).Visible = False
.Columns(16).Visible = False
.Columns(17).Visible = True
.Columns(18).Visible = True
.Columns(19).Visible = False
.Columns(20).Visible = False
.Columns(21).Visible = True
.Columns(22).Visible = True
.Columns(23).Visible = False
.Columns(24).Visible = False
.Columns(25).Visible = False
.Columns(26).Visible = False
.Columns(27).Visible = False
.Columns(28).Visible = False
.Columns(29).Visible = False
.Columns(30).Visible = False
.Columns(31).Visible = False
.Columns(32).Visible = False
.Columns(33).Visible = False
.Columns(34).Visible = False
.Columns(35).Visible = False
.Columns(36).Visible = True
.Columns(37).Visible = False
.Columns(38).Visible = False
.Columns(39).Visible = True
.Columns(40).Visible = True
.Columns(41).Visible = False
.Columns(42).Visible = False
.Columns(43).Visible = False
.Columns(44).Visible = False
.Columns(45).Visible = False
.Columns(46).Visible = False
End With

Set DataGrid3.DataSource = rs5
With DataGrid3
.Columns(0).Visible = False
.Columns(1).Visible = False
.Columns(2).Visible = True
.Columns(3).Visible = False
.Columns(4).Visible = True
.Columns(5).Visible = False
.Columns(6).Visible = False
.Columns(7).Visible = False
.Columns(8).Visible = False
.Columns(9).Visible = False
.Columns(10).Visible = False
.Columns(11).Visible = False
.Columns(12).Visible = False
.Columns(13).Visible = False
.Columns(14).Visible = True
.Columns(15).Visible = False
.Columns(16).Visible = False
.Columns(17).Visible = False
.Columns(18).Visible = False
.Columns(19).Visible = False
.Columns(20).Visible = False
.Columns(21).Visible = False
.Columns(22).Visible = False
.Columns(23).Visible = False
.Columns(24).Visible = False
.Columns(25).Visible = True
.Columns(26).Visible = True
.Columns(27).Visible = False
.Columns(28).Visible = False
.Columns(29).Visible = True
.Columns(30).Visible = True
.Columns(31).Visible = False
.Columns(32).Visible = False
.Columns(33).Visible = False
.Columns(34).Visible = False
.Columns(35).Visible = False
.Columns(36).Visible = True
.Columns(37).Visible = False
.Columns(38).Visible = False
.Columns(39).Visible = True
.Columns(40).Visible = True
.Columns(41).Visible = False
.Columns(42).Visible = False
.Columns(43).Visible = False
.Columns(44).Visible = False
.Columns(45).Visible = False
.Columns(46).Visible = False

End With

End Sub
Private Sub SimpanLKO()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
'Set DataGrid2.DataSource = rs2
End Sub

 Private Sub Form_Resize()
On Error GoTo err
If Me.Width >= (1366 * Screen.TwipsPerPixelX) Or Me.Width <= (1366 * Screen.TwipsPerPixelX) Then
   Me.Width = (NoMesin.Text10 * Screen.TwipsPerPixelX)
   End If
If Me.Height >= (768 * Screen.TwipsPerPixelX) Or Me.Height <= (768 * Screen.TwipsPerPixelX) Then
   Me.Height = (NoMesin.Text11 * Screen.TwipsPerPixelX)
   End If
Exit Sub
err:
    Me.WindowState = 0
End Sub


Private Sub Label42_Change()
'Call MuatUlang
Label43.Caption = Label42

End Sub

Private Sub Label43_Change()
'Command6_Click

End Sub

Private Sub MSComm1_OnComm()
If Len(Text6.Text) > 0 Then
        MSComm1.Output = Text6.Text
        Text6.Text = vbNullString
    End If
    Text6.SetFocus

End Sub

Private Sub Text11_DblClick()
Text11.Text = Format(Now, "yyyy/mm/dd")
End Sub

Private Sub Text16_Change()
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
End If
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

Private Sub Text28_Change()
Dim hitung As Integer
hitung = Len(Text1.Text)
If hitung = 4 Then
Call Koneksi
End If
End Sub

Private Sub Timer1_Timer()
    Label30.Caption = time
    Label42.Caption = Format(Now, "hh.mm")
End Sub


Public Sub NIKOP()

Set jon = New ADODB.Connection
Set rs4 = New ADODB.Recordset
jon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs4.Open "select*from NIKOP", jon, adOpenDynamic, adLockOptimistic
Do While Not rs4.EOF
Combo4.AddItem rs4!OPERATOR
rs4.MoveNext

Loop
End Sub

Private Sub NIKRE()
Set Con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
Con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

rs2.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
'Combo4.Text = rs2!NIK
Combo2.Text = rs2!Shift
End Sub

Private Sub HasilSimpan()

Set rs3 = New ADODB.Recordset
rs3.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
Set DataGrid4.DataSource = rs3
rs3.Sort = DataGrid4.Columns(34).DataField & " DESC"

With DataGrid4
.Columns(0).Visible = False
.Columns(1).Visible = False
.Columns(2).Visible = False
.Columns(3).Visible = False
.Columns(4).Visible = False
.Columns(5).Visible = False
.Columns(6).Visible = True
.Columns(7).Visible = True
.Columns(6).Width = 36
.Columns(7).Width = 36

End With

End Sub

Sub WaktuTunggu()
Dim time1, time2

time1 = Now
time2 = Now + TimeValue("0:00:03")
    Do Until time1 >= time2
        DoEvents
        time1 = Now()
    Loop

End Sub
Public Sub Koneksi()
Kon2.Text2.Text = Label45 + Text1.Text
Kon2.Text3.Text = Text1.Text
Kon2.Text1.Text = Text26.Text
Kon2.Text4.Text = Text28.Text
End Sub

Public Sub MuatUlang()

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus

Unload Me
'Unload Koneksi

'WaktuTunggu

'Koneksi.Show

FrmOP.Show

Text28.Text = "0"
Text28.Text = DataGrid1.Columns(1)

Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

'Call NIKRE

End Sub

Public Sub Abjad()
If Combo5 = "01" Then
Label45.Caption = "A"
End If

If Combo5 = "02" Then
Label45.Caption = "B"
End If

If Combo5 = "03" Then
Label45.Caption = "C"
End If

If Combo5 = "04" Then
Label45.Caption = "D"
End If

If Combo5 = "05" Then
Label45.Caption = "E"
End If

If Combo5 = "06" Then
Label45.Caption = "F"
End If

If Combo5 = "07" Then
Label45.Caption = "G"
End If

If Combo5 = "08" Then
Label45.Caption = "H"
End If

If Combo5 = "09" Then
Label45.Caption = "I"
End If

If Combo5 = "10" Then
Label45.Caption = "J"
End If

If Combo5 = "11" Then
Label45.Caption = "K"
End If

If Combo5 = "12" Then
Label45.Caption = "L"
End If

If Combo5 = "13" Then
Label45.Caption = "M"
End If

If Combo5 = "14" Then
Label45.Caption = "N"
End If

If Combo5 = "15" Then
Label45.Caption = "O"
End If

If Combo5 = "16" Then
Label45.Caption = "P"
End If

If Combo5 = "17" Then
Label45.Caption = "Q"
End If

If Combo5 = "18" Then
Label45.Caption = "R"
End If

If Combo5 = "19" Then
Label45.Caption = "S"
End If

If Combo5 = "20" Then
Label45.Caption = "T"
End If

If Combo5 = "21" Then
Label45.Caption = "U"
End If

If Combo5 = "22" Then
Label45.Caption = "V"
End If

If Combo5 = "23" Then
Label45.Caption = "W"
End If

If Combo5 = "24" Then
Label45.Caption = "X"
End If

If Combo5 = "25" Then
Label45.Caption = "Y"
End If

If Combo5 = "26" Then
Label45.Caption = "Z"
End If

If Combo5 = "27" Then
Label45.Caption = "A"
End If

If Combo5 = "28" Then
Label45.Caption = "B"
End If

If Combo5 = "29" Then
Label45.Caption = "C"
End If

If Combo5 = "30" Then
Label45.Caption = "D"
End If

End Sub
