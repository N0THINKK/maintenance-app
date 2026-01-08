VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   9285
   ClientLeft      =   240
   ClientTop       =   885
   ClientWidth     =   9300
   ControlBox      =   0   'False
   Icon            =   "FrmLKO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   619
   ScaleMode       =   0  'User
   ScaleWidth      =   620
   Begin VB.TextBox Text37 
      Height          =   285
      Left            =   360
      TabIndex        =   109
      Text            =   "B"
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text36 
      Height          =   285
      Left            =   120
      TabIndex        =   108
      Text            =   "A"
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "TimeOn"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8280
      TabIndex        =   107
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Pilih Otomatis"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   106
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   1320
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Height          =   1935
      Left            =   7680
      TabIndex        =   104
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   14
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   2415
      Left            =   7680
      TabIndex        =   103
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   14
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1575
      Left            =   4800
      TabIndex        =   102
      Top             =   7680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
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
      Height          =   1575
      Left            =   120
      TabIndex        =   101
      Top             =   7680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
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
      Height          =   4695
      Left            =   6120
      TabIndex        =   100
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   14
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
   Begin VB.CheckBox Check3 
      Caption         =   "Simpan Otomatis"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   97
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text35 
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
      Left            =   3000
      TabIndex        =   94
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Track Defect"
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
      TabIndex        =   93
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text34 
      Height          =   285
      Left            =   8040
      TabIndex        =   91
      Text            =   "Text34"
      Top             =   8760
      Width           =   735
   End
   Begin VB.TextBox Text33 
      Height          =   285
      Left            =   8160
      TabIndex        =   90
      Text            =   "Text33"
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox Text32 
      Height          =   285
      Left            =   8160
      TabIndex        =   89
      Text            =   "Text32"
      Top             =   8640
      Width           =   735
   End
   Begin VB.TextBox Text31 
      Height          =   285
      Left            =   8160
      TabIndex        =   88
      Text            =   "Text31"
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   8160
      TabIndex        =   87
      Text            =   "Text30"
      Top             =   8160
      Width           =   735
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   8160
      TabIndex        =   86
      Text            =   "Text29"
      Top             =   7920
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5760
      Top             =   1800
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   6480
      TabIndex        =   82
      Text            =   "Text26"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   6480
      TabIndex        =   81
      Text            =   "Text25"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   240
      TabIndex        =   74
      Top             =   6900
      Width           =   255
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
      TabIndex        =   72
      Text            =   "0"
      Top             =   6480
      Width           =   1215
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
      TabIndex        =   65
      Top             =   2520
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
      TabIndex        =   64
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   3480
      Top             =   8040
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
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7440
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
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   120
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
      Height          =   615
      Left            =   4800
      TabIndex        =   58
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
      Left            =   5520
      TabIndex        =   52
      Top             =   840
      Width           =   1575
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
      TabIndex        =   50
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FrmLKO.frx":1084A
      Left            =   1440
      List            =   "FrmLKO.frx":1084C
      TabIndex        =   46
      Top             =   5760
      Width           =   3375
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
      TabIndex        =   44
      Top             =   1200
      Width           =   1455
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
      TabIndex        =   42
      Top             =   4200
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
      Left            =   1440
      TabIndex        =   41
      Text            =   "0"
      Top             =   6480
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
      Left            =   4800
      TabIndex        =   39
      Text            =   "0"
      Top             =   5040
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
      Left            =   4800
      TabIndex        =   38
      Text            =   "0"
      Top             =   4680
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
      ItemData        =   "FrmLKO.frx":1084E
      Left            =   4080
      List            =   "FrmLKO.frx":1085B
      TabIndex        =   32
      Top             =   840
      Width           =   1335
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
      Height          =   855
      Left            =   240
      TabIndex        =   31
      Top             =   5760
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
      ItemData        =   "FrmLKO.frx":10869
      Left            =   1560
      List            =   "FrmLKO.frx":10888
      TabIndex        =   29
      Top             =   840
      Width           =   1335
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
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   1335
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
      Left            =   4800
      TabIndex        =   25
      Text            =   "0"
      Top             =   4320
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
      Left            =   4800
      TabIndex        =   23
      Text            =   "0"
      Top             =   3960
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
      Left            =   4800
      TabIndex        =   21
      Top             =   3480
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
      TabIndex        =   20
      Text            =   "0"
      Top             =   5040
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
      TabIndex        =   19
      Text            =   "0"
      Top             =   4680
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
      TabIndex        =   18
      Text            =   "0"
      Top             =   4320
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
      TabIndex        =   17
      Text            =   "0"
      Top             =   3960
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
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SIMPAN"
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
      Top             =   5760
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
      Top             =   2520
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
      Top             =   3360
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6720
      Top             =   5040
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
      TabIndex        =   57
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
      TabIndex        =   56
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
      Left            =   3720
      TabIndex        =   48
      Top             =   1800
      Width           =   1335
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
      TabIndex        =   60
      Top             =   8400
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
      TabIndex        =   61
      Top             =   8400
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
      Left            =   6480
      TabIndex        =   45
      Top             =   1200
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
      Left            =   7440
      TabIndex        =   53
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text28 
      Height          =   315
      Left            =   3480
      TabIndex        =   80
      Text            =   "Text28"
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label51 
      Caption         =   "Angka"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   105
      Top             =   2760
      Width           =   495
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
      Left            =   7920
      TabIndex        =   99
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
      TabIndex        =   98
      Top             =   1800
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      Height          =   735
      Left            =   2880
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label50 
      BackColor       =   &H0000FFFF&
      Caption         =   "Pastikan No Urut sesuai dengan Kanban asal nomer mesin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   96
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label49 
      BackColor       =   &H0000FFFF&
      Caption         =   "Pastikan Sequen Stripping no.2 disimpan dahulu. Abaikan jika sudah"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   95
      Top             =   5400
      Width           =   6015
   End
   Begin VB.Label Label48 
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
      Left            =   8280
      TabIndex        =   92
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label47 
      Caption         =   "BarcodeKanban"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7920
      TabIndex        =   85
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label46 
      Caption         =   "Label46"
      Height          =   255
      Left            =   5520
      TabIndex        =   84
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label45 
      Caption         =   "Huruf"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   83
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label40 
      Height          =   255
      Left            =   1800
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label39 
      Height          =   255
      Left            =   1560
      TabIndex        =   78
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label42 
      Height          =   255
      Left            =   240
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label43 
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "FrmLKO.frx":108DF
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label41 
      Caption         =   "NoLogin"
      Height          =   255
      Left            =   600
      TabIndex        =   75
      Top             =   6840
      Width           =   495
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
      Left            =   3480
      TabIndex        =   73
      Top             =   6240
      Width           =   1215
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
      Left            =   3000
      TabIndex        =   71
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label37 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   6360
      TabIndex        =   70
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label38 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   6720
      TabIndex        =   69
      Top             =   120
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   68
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
      Left            =   6480
      TabIndex        =   67
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
      TabIndex        =   66
      Top             =   2535
      Width           =   750
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
      TabIndex        =   59
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label28 
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
      Left            =   7440
      TabIndex        =   54
      Top             =   840
      Width           =   735
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
      Left            =   8040
      TabIndex        =   55
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
      TabIndex        =   51
      Top             =   4800
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
      Left            =   3000
      TabIndex        =   49
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
      Left            =   1560
      TabIndex        =   47
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Picture         =   "FrmLKO.frx":10D4B
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
      TabIndex        =   43
      Top             =   3120
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
      Left            =   3840
      TabIndex        =   37
      Top             =   3960
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
      Left            =   3840
      TabIndex        =   36
      Top             =   4680
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
      Left            =   3840
      TabIndex        =   35
      Top             =   4320
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
      Left            =   3840
      TabIndex        =   34
      Top             =   5040
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
      Left            =   4080
      TabIndex        =   33
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
      Left            =   1560
      TabIndex        =   30
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
      Left            =   5520
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   24
      Top             =   3960
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
      Left            =   1440
      TabIndex        =   22
      Top             =   6240
      Width           =   2055
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
      TabIndex        =   15
      Top             =   5040
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
      TabIndex        =   14
      Top             =   4320
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
      TabIndex        =   13
      Top             =   4680
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
      TabIndex        =   12
      Top             =   3960
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
      Left            =   3840
      TabIndex        =   11
      Top             =   3480
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
      TabIndex        =   10
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "SISI B"
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
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "SISI A"
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
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      Top             =   2520
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
      TabIndex        =   40
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FrmOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Public rs6 As New ADODB.Recordset

Public con As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset

Public jon As New ADODB.Connection
Public rs4 As New ADODB.Recordset

Dim total
Dim total1
Dim TeA
Dim TeB

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
CrimpingStandart
Else
CrimpStandartFull.Hide
End If

Label46 = 0

End Sub

Private Sub Combo3_Change()
Label46 = 0
End Sub


Private Sub Command9_Click()
If Check4 = 1 Then
Timer4.Enabled = True
End If

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

Label46 = 0
End Sub

Private Sub DataGrid4_Click()
Label46 = 0
End Sub

Private Sub Label30_Change()
    
    If Label49.BackColor = vbYellow Then
    Label49.BackColor = vbBlue
    Label49.ForeColor = vbWhite
    ElseIf Label49.BackColor = vbBlue Then
    Label49.BackColor = vbYellow
    Label49.ForeColor = vbBlack
    End If
    
    If Label50.BackColor = vbYellow Then
    Label50.BackColor = vbBlue
    Label50.ForeColor = vbWhite
    ElseIf Label50.BackColor = vbBlue Then
    Label50.BackColor = vbYellow
    Label50.ForeColor = vbBlack
    End If
    
    If Shape1.BorderColor = vbYellow Then
    Shape1.BorderColor = vbBlue
    ElseIf Shape1.BorderColor = vbBlue Then
    Shape1.BorderColor = vbYellow
    End If
    
    If Check3 = 1 And Command4.Enabled = True Then
   ' Command9.Caption = "AUTO"
    Else
   ' Command9.Caption = "Test"

    End If
    
End Sub

Private Sub Text1_Change()
TampilUkur
End Sub

Private Sub Text2_Change()
Label46 = 0
End Sub

Private Sub Text20_Change()
tampil
End Sub

Private Sub Text21_Change()
Label46 = 0
End Sub

Private Sub Text3_Change()
Label46 = 0
End Sub

Private Sub text35_Change()
Abjad
Text28 = ""
Text28.Text = DataGrid1.Columns(1)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Kon2.Show
End Sub

Private Sub Command4_Click()

Call SimpanData

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

Unload Me

FrmOP.Show

Text28.Text = ""

Text35.Text = DataGrid4.Columns(35)
Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

'Label25.Caption = DataLKO.DataGrid1.Columns(8)
'Text18.Text = DataLKO.DataGrid1.Columns(9)
'Text19.Text = DataLKO.DataGrid1.Columns(11)
'Text20.Text = DataLKO.DataGrid1.Columns(6)
'Label5.Caption = DataLKO.DataGrid1.Columns(10)
'Label6.Caption = DataLKO.DataGrid1.Columns(12)

Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

'Call NIKRE
Label48 = total1
Label28 = total
Label31 = TeA
Label32 = TeB

Label49.Visible = False

'Call WaktuTunggu
Shell App.Path & "\SimpanLKO.bat", vbMinimizedFocus

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus
End Sub

Private Sub Command5_Click()
DataLKOOP.Show
End Sub

Private Sub Command6_Click()

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus

Unload CrimpStandart
Unload CrimpStandartFull

Unload Me

WaktuTunggu

FrmOP.Show

Text28.Text = ""

Text35.Text = DataGrid4.Columns(35)
Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

Jumlah
Label48 = total1
Label28 = total
Label31 = TeA
Label32 = TeB

Label49.Visible = False

End Sub

Private Sub Command8_Click()
TrackDefect.Show
End Sub

Private Sub DataGrid1_Click()

Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

Label46 = 0

Text31.Text = Label5
Text33.Text = Label25
Text32.Text = Label6
Text29.Text = Text18
Text30.Text = Text19
Text34.Text = Label51

'Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(6)
'Label25.Caption = DataGrid1.Columns(42)
'Text18.Text = DataGrid1.Columns(10)
'Text19.Text = DataGrid1.Columns(11)
'Text20.Text = DataGrid1.Columns(4)
'Label5.Caption = DataGrid1.Columns(14)
'Label6.Caption = DataGrid1.Columns(15)
'Text12.Text = DataGrid1.Columns(1)
'Text22.Text = DataGrid1.Columns(2)
'Text24.Text = DataGrid1.Columns(41)

'CrimpStandart.Text3.Text = DataGrid1.Columns(42)
'CrimpStandart.Text1.Text = DataGrid1.Columns(10)
'CrimpStandart.Text2.Text = DataGrid1.Columns(11)

Call Otomatis

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

Private Sub DataGrid5_Click()
Text31.Text = DataGrid5.Columns(14)
Text33.Text = DataGrid5.Columns(42)
Text32.Text = DataGrid5.Columns(15)
Text29.Text = DataGrid5.Columns(10)
Text30.Text = DataGrid5.Columns(11)
Text34.Text = DataGrid5.Columns(0)

TrackDefect.Text18 = DataGrid5.Columns(10)
TrackDefect.Label16 = DataGrid5.Columns(14)
TrackDefect.Text19 = DataGrid5.Columns(11)
TrackDefect.Label15 = DataGrid5.Columns(15)
TrackDefect.Text5 = DataGrid5.Columns(0)
TrackDefect.Label25 = DataGrid5.Columns(42)

CrimpStandart.Text3 = Text33
CrimpStandart.Text4 = Text31
CrimpStandart.Text5 = Text32
CrimpStandart.Text1 = Text29
CrimpStandart.Text2 = Text30

CrimpStandartFull.Text3 = Text33
CrimpStandartFull.Text4 = Text31
CrimpStandartFull.Text5 = Text32
CrimpStandartFull.Text1 = Text29
CrimpStandartFull.Text2 = Text30

Label46 = 0

End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Jissk.mdb"

Set con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

'Combo1.Text = NoMesin.Text21.Text

Combo2.Text = StrConv(Combo2.Text, vbProperCase)

Dim henti As Boolean
henti = False
'Do Until henti
Text11 = Format(Now, "yyyy/mm/dd")
Label37 = Format(Now, "yyyy")
Label38 = Format(Now, "mm")
'Label41 = FrmUtama.Label6

DoEvents
'Loop

Text13.Text = DataGrid2.Columns(17)
Text23.Text = DataGrid3.Columns(25)

'Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(6)
'Label25.Caption = DataGrid1.Columns(42)
'Text18.Text = DataGrid1.Columns(10)
'Text19.Text = DataGrid1.Columns(11)
'Text20.Text = DataGrid1.Columns(4)
'Label5.Caption = DataGrid1.Columns(14)
'Label6.Caption = DataGrid1.Columns(15)
'Text12.Text = DataGrid1.Columns(1)
'Text22.Text = DataGrid1.Columns(2)

'Text4.Text = DataGrid2.Columns(17)
'Text5.Text = DataGrid2.Columns(18)
'Text6.Text = DataGrid2.Columns(21)
'Text7.Text = DataGrid2.Columns(22)
'Text9.Text = DataGrid3.Columns(25)
'Text10.Text = DataGrid3.Columns(26)
'Text14.Text = DataGrid3.Columns(29)
'Text15.Text = DataGrid3.Columns(30)

'TampilUkur

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

'Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(4)
'Text25.Text = DataGrid1.Columns(2)
'Text26.Text = DataGrid1.Columns(3)
'Text27.Text = DataGrid1.Columns(5)
'Text28.Text = DataGrid1.Columns(1)

'Label48 = total1
'Label28 = total
'Label31 = TeA
'Label32 = TeB


End Sub

Private Sub Form_Load()

Me.Top = NoMesin.Text8
Me.Left = NoMesin.Text6

Form_Resize

Label49.BackColor = vbYellow

Label50.BackColor = vbYellow

Shape1.BorderColor = vbYellow

Text28.Text = ""

Set kon = New ADODB.Connection
kon.CursorLocation = adUseClient
kon.Provider = "Microsoft.jet.oledb.4.0"
kon.Open App.Path & "\Jissk.mdb"
Call LKO

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\DbLKO.mdb"
Call SimpanLKO
'rs2.Sort = DataGrid2.Columns(0).DataField & " DESC"

Set Bon = New ADODB.Connection
Bon.CursorLocation = adUseClient
Bon.Provider = "Microsoft.jet.oledb.4.0"
Bon.Open App.Path & "\DBchcw.mdb"

'Text1 = DataGrid1.Columns(0)
'Text17 = DataGrid1.Columns(6)
'Label25.Caption = DataGrid1.Columns(42)
'Text18.Text = DataGrid1.Columns(10)
'Text19.Text = DataGrid1.Columns(11)
'Text20.Text = DataGrid1.Columns(4)
'Label5.Caption = DataGrid1.Columns(14)
'Label6.Caption = DataGrid1.Columns(15)
'Text12.Text = DataGrid1.Columns(1)
'Text22.Text = DataGrid1.Columns(2)

With Combo3
.AddItem "A.1 Core Terurai"
.AddItem "A.2 Core Terpotong"
.AddItem "A.3 Core Tidak teratur"
.AddItem "A.4 Core Maju"
.AddItem "A.5 Core Mundur"
.AddItem "A.6 Tidak Tercrimping"
.AddItem "A.7 Scracth"
.AddItem "B.1 Terminal Tergores"
.AddItem "B.2 Terminal Bengkok ke atas"
.AddItem "B.3 Terminal Bengkok ke bawah"
.AddItem "B.4 Terminal Melintir"
.AddItem "B.5 Terminal Ujung Terpotong"
.AddItem "B.6 Terminal Ujung Terbuka"
.AddItem "B.7 Terminal Ujung Rusak"
.AddItem "B.8 Terminal Bridge terlalu panjang"
.AddItem "B.9 Terminal Centilever Rusak"
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
.AddItem "D.6 Tidak Standart"
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
.AddItem "H.4 Kondisi core bagian A"
.AddItem "H.5 Resin masuk bagian A"
.AddItem "H.6 Resin barel bagian B Terbuka"
.AddItem "H.7 Core terlihat atas sisi C"
.AddItem "H.8 Core terlihat samping sisi C"
.AddItem "H.9 Sisi punggung"
.AddItem "H.10 Abnormal resin"
.AddItem "H.11 Panjang welding N-OK"
.AddItem "H.12 Circuit tidak terbonder"
.AddItem "H.13 Bonder Retak"
.AddItem "H.14 Stripping kepanjangan"
End With

Call HasilSimpan
rs3.Sort = DataGrid4.Columns(36).DataField & " DESC"

'Text35.Text = DataGrid4.Columns(35)

Text28.Text = ""

Text34.Text = DataGrid5.Columns(0)
Text31.Text = DataGrid5.Columns(14)
Text33.Text = DataGrid5.Columns(42)
Text32.Text = DataGrid5.Columns(15)
Text29.Text = DataGrid5.Columns(10)
Text30.Text = DataGrid5.Columns(11)

Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)

Combo1.Text = NoMesin.Text21.Text

Label41.Caption = DataGrid4.Columns(34)

Text35.Text = Right$(Combo1.Text, 2)

Text28.Text = DataGrid1.Columns(1)

Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

'Do Until Command4.Enabled = False
'rs1.MoveNext
'Loop

Jumlah

Text36 = DataGrid2.Columns(2)
Text37 = DataGrid3.Columns(2)

End Sub

Private Sub LKO()
'Set rs1 = New ADODB.Recordset
'rs1.Open "select*from ProductAC81", kon, adOpenDynamic, adLockOptimistic
'Set DataGrid1.DataSource = rs1
'rs1.Sort = DataGrid1.Columns(55).DataField & " DESC"

Set rs1 = New ADODB.Recordset
rs1.Open "select*from Prdlog", kon, adOpenForwardOnly, adLockReadOnly
Set DataGrid1.DataSource = rs1
rs1.Sort = DataGrid1.Columns(3).DataField & " DESC"

Set rs5 = New ADODB.Recordset
rs5.Open "select*from Hasil", kon, adOpenForwardOnly, adLockReadOnly
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

Set rs6 = New ADODB.Recordset
rs6.Open "select distinct*from Product", kon, adOpenForwardOnly, adLockReadOnly
Set DataGrid5.DataSource = rs6
rs6.Sort = DataGrid5.Columns(36).DataField & " ASC"

'rs1.Close

End Sub
Private Sub SimpanLKO()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from LKO", con, adOpenForwardOnly, adLockReadOnly
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


Private Sub MSComm1_OnComm()
If Len(Text6.Text) > 0 Then
        MSComm1.Output = Text6.Text
        Text6.Text = vbNullString
    End If
    Text6.SetFocus

End Sub

Private Sub Text10_Change()
Label46 = 0
End Sub

Private Sub Text11_DblClick()
Text11.Text = Format(Now, "yyyy/mm/dd")
End Sub

Private Sub Text14_Change()
Label46 = 0
End Sub

Private Sub text15_Change()
Label46 = 0
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

AngkaLengkap
Dim hitung As Integer
hitung = Len(Text1.Text)
If hitung >= 1 Then
Call Koneksi
End If

If Command4.Enabled = True And Check3 = 1 Then
Call SimpanData
End If

End Sub

Private Sub Text29_Change()
CrimpStandart.Text3.Text = Text33.Text
CrimpStandart.Text4.Text = Text31.Text
CrimpStandart.Text5.Text = Text32.Text
CrimpStandart.Text1.Text = Text29.Text
CrimpStandart.Text2.Text = Text30.Text

CrimpStandartFull.Text3.Text = Text33.Text
CrimpStandartFull.Text4.Text = Text31.Text
CrimpStandartFull.Text5.Text = Text32.Text
CrimpStandartFull.Text1.Text = Text29.Text
CrimpStandartFull.Text2.Text = Text30.Text
End Sub

Private Sub Text30_Change()
CrimpStandart.Text3 = Text33.Text
CrimpStandart.Text4 = Text31.Text
CrimpStandart.Text5 = Text32.Text
CrimpStandart.Text1 = Text29.Text
CrimpStandart.Text2 = Text30.Text

CrimpStandartFull.Text3 = Text33.Text
CrimpStandartFull.Text4 = Text31.Text
CrimpStandartFull.Text5 = Text32.Text
CrimpStandartFull.Text1 = Text29.Text
CrimpStandartFull.Text2 = Text30.Text
End Sub

Private Sub Text4_Change()
Label46 = 0
End Sub

Private Sub Text5_Change()
Label46 = 0
End Sub

Private Sub Text6_Change()
Label46 = 0
End Sub

Private Sub Text7_Change()
Label46 = 0
End Sub

Private Sub Text8_Change()
Label46 = 0
End Sub

Private Sub Text9_Change()
Label46 = 0
End Sub

Private Sub Timer1_Timer()
    'lbltgl.Caption = Format(Date, "dddd, d mmmm yyyy")
    Label30.Caption = time
    
End Sub

Public Sub NIKOP()

Set jon = New ADODB.Connection
Set rs4 = New ADODB.Recordset
jon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs4.Open "select*from NIKOP", jon, adOpenForwardOnly, adLockReadOnly
Do While Not rs4.EOF
Combo4.AddItem rs4!OPERATOR
'rs4.MoveNext

Loop
End Sub

Private Sub NIKRE()
Set con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbLKO.mdb"

rs2.Open "select*from LKO", con, adOpenForwardOnly, adLockReadOnly
Combo4.Text = rs2!NIK
Combo2.Text = rs2!Shift
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

'Sub loaddata()

'Dim ch As ColumnHeader
  
'Set ch = ListView1.ColumnHeaders.Add(, , "Sequen", 1700)
'Set ch = ListView1.ColumnHeaders.Add(, , "Size", 2100, vbLeftJustify)

'Dim list As ListItem
'ListView1.ListItems.Clear

'If rs3.State = 1 Then rs3.Close
'rs3.Open "select*from LKO", Con, adOpenDynamic, adLockOptimistic
'While Not rs3.EOF
'Do Until rs3.EOF
'Set list = ListView1.ListItems.Add(, , rs3.Fields(28))
'list.SubItems(1) = rs3.Fields(4)
'list.SubItems(2) = rs3.Fields(2)
'list.SubItems(3) = rs3.Fields(3)
'list.SubItems(4) = rs3.Fields(4)
'list.SubItems(5) = rs3.Fields(5)
'list.SubItems(6) = rs3.Fields(6)
'list.SubItems(7) = rs3.Fields(7)
'list.SubItems(8) = rs3.Fields(8)
'list.SubItems(9) = rs3.Fields(9)
'list.SubItems(10) = rs3.Fields(10)
'list.SubItems(11) = rs3.Fields(11)
'list.SubItems(12) = rs3.Fields(12)
'list.SubItems(13) = rs3.Fields(13)
'list.SubItems(14) = rs3.Fields(14)
'list.SubItems(15) = rs3.Fields(15)
'list.SubItems(16) = rs3.Fields(16)
'list.SubItems(17) = rs3.Fields(17)
'list.SubItems(18) = rs3.Fields(18)
'list.SubItems(19) = rs3.Fields(19)
'list.SubItems(20) = rs3.Fields(20)
'list.SubItems(21) = rs3.Fields(21)
'list.SubItems(22) = rs3.Fields(22)
'list.SubItems(23) = rs3.Fields(23)
'list.SubItems(24) = rs3.Fields(24)
'list.SubItems(25) = rs3.Fields(25)
'list.SubItems(26) = rs3.Fields(26)
'list.SubItems(27) = rs3.Fields(27)
'list.SubItems(28) = rs3.Fields(28)
'list.SubItems(29) = rs3.Fields(29)
'list.SubItems(30) = rs3.Fields(30)
'list.SubItems(31) = rs3.Fields(31)
'list.SubItems(32) = rs3.Fields(32)
'list.SubItems(33) = rs3.Fields(33)
'list.SubItems(34) = rs3.Fields(34)
'list.SubItems(35) = rs3.Fields(35)
'list.SubItems(36) = rs3.Fields(36)
'list.SubItems(37) = rs3.Fields(37)
'list.SubItems(38) = rs3.Fields(38)
'list.SubItems(39) = rs3.Fields(39)
'list.SubItems(40) = rs3.Fields(40)
'list.SubItems(41) = rs3.Fields(41)
'list.SubItems(42) = rs3.Fields(42)
'list.SubItems(43) = rs3.Fields(43)
'list.SubItems(44) = rs3.Fields(44)
'list.SubItems(45) = rs3.Fields(45)
'list.SubItems(46) = rs3.Fields(46)
'list.SubItems(47) = rs3.Fields(47)
'list.SubItems(48) = rs3.Fields(48)
'list.SubItems(49) = rs3.Fields(49)
'list.SubItems(50) = rs3.Fields(50)
'list.SubItems(51) = rs3.Fields(51)
'list.SubItems(52) = rs3.Fields(52)
'list.SubItems(53) = rs3.Fields(53)
'list.SubItems(54) = rs3.Fields(54)
'list.SubItems(55) = rs3.Fields(55)
'list.SubItems(56) = rs3.Fields(56)


'rs3.MoveNext
'Loop

'Wend
'End Sub

Private Sub HasilSimpan()
Set rs3 = New ADODB.Recordset
rs3.Open "SELECT TOP 500 * FROM LKO", con, adOpenForwardOnly, adLockReadOnly
'rs3.Open "SELECT TOP 500 * FROM LKO ORDER BY ScanTime DESC", con, adOpenForwardOnly, adLockReadOnly
'rs3.Open "select*from LKO", con, adOpenDynamic, adLockOptimistic
Set DataGrid4.DataSource = rs3
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

End With

End Sub

Public Sub Koneksi()
Kon2.Text2.Text = Label45 + Label51
Kon2.Text3.Text = Text1.Text
Kon2.Text1.Text = Text26.Text
Kon2.Text4.Text = Text28.Text
End Sub

Public Sub Abjad()

If Text35 = "" Then
Label45.Caption = ""
End If

If Text35 = "01" Then
Label45.Caption = "A"
End If

If Text35 = "02" Then
Label45.Caption = "B"
End If

If Text35 = "03" Then
Label45.Caption = "C"
End If

If Text35 = "04" Then
Label45.Caption = "D"
End If

If Text35 = "05" Then
Label45.Caption = "E"
End If

If Text35 = "06" Then
Label45.Caption = "F"
End If

If Text35 = "07" Then
Label45.Caption = "G"
End If

If Text35 = "08" Then
Label45.Caption = "H"
End If

If Text35 = "09" Then
Label45.Caption = "I"
End If

If Text35 = "10" Then
Label45.Caption = "J"
End If

If Text35 = "11" Then
Label45.Caption = "K"
End If

If Text35 = "12" Then
Label45.Caption = "L"
End If

If Text35 = "13" Then
Label45.Caption = "M"
End If

If Text35 = "14" Then
Label45.Caption = "N"
End If

If Text35 = "15" Then
Label45.Caption = "O"
End If

If Text35 = "16" Then
Label45.Caption = "P"
End If

If Text35 = "17" Then
Label45.Caption = "Q"
End If

If Text35 = "18" Then
Label45.Caption = "R"
End If

If Text35 = "19" Then
Label45.Caption = "S"
End If

If Text35 = "20" Then
Label45.Caption = "T"
End If

If Text35 = "21" Then
Label45.Caption = "U"
End If

If Text35 = "22" Then
Label45.Caption = "V"
End If

If Text35 = "23" Then
Label45.Caption = "W"
End If

If Text35 = "24" Then
Label45.Caption = "X"
End If

If Text35 = "25" Then
Label45.Caption = "Y"
End If

If Text35 = "26" Then
Label45.Caption = "Z"
End If

If Text35 = "27" Then
Label45.Caption = "A"
End If

If Text35 = "28" Then
Label45.Caption = "B"
End If

If Text35 = "29" Then
Label45.Caption = "C"
End If

If Text35 = "30" Then
Label45.Caption = "D"
End If

If Text35 = "31" Then
Label45.Caption = "E"
End If

If Text35 = "32" Then
Label45.Caption = "F"
End If

If Text35 = "33" Then
Label45.Caption = "G"
End If

If Text35 = "34" Then
Label45.Caption = "H"
End If

If Text35 = "35" Then
Label45.Caption = "I"
End If

If Text35 = "36" Then
Label45.Caption = "J"
End If

If Text35 = "37" Then
Label45.Caption = "K"
End If

If Text35 = "38" Then
Label45.Caption = "L"
End If


End Sub

Public Sub hitung()
Label46.Caption = Val(Label46) + 1

If Label46 = 60 Then
BukaUlang

Label46 = 0
End If

End Sub

Private Sub Timer3_Timer()
If Combo2 = "" Then
Label46 = 0
Command6.Enabled = False
ElseIf Combo2 <> "" Then
Call hitung
Command6.Enabled = True

End If

End Sub

Private Sub Jumlah()

Form_Activate
rs2.Open "select * from LKO where Tanggal = '" & Text11 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'", con
'rs2.MoveFirst
total = 0
Do While Not rs2.EOF
    total = total + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and Shift = '" & Combo2 & "'", con
'rs2.MoveFirst
total1 = 0
Do While Not rs2.EOF
    total1 = total1 + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and Shift = '" & Combo2 & "' and TermB = '" & Text19 & "'", con
'rs2.MoveFirst
TeB = 0
Do While Not rs2.EOF
    TeB = TeB + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

rs2.Open "select * from LKO where NoLogin = '" & Label41 & "' and Shift = '" & Combo2 & "' and TermA = '" & Text18 & "'", con
'rs2.MoveFirst
TeA = 0
Do While Not rs2.EOF
    TeA = TeA + rs2!QtyProduct
   rs2.MoveNext
Loop
rs2.Close

End Sub

Private Sub KriteriaHitung()
Dim Masuk
Masuk = Label41.Caption
If Text28 = "2" Then
Masuk = Masuk + 1

Label41.Caption = Masuk
End If
End Sub

Private Sub CrimpingStandart()

If NoMesin.Text14 <> "" And NoMesin.Text15 <> 0 Then
CrimpStandartFull.Show
End If

End Sub

Private Sub tampil()

If Text20.Text = "" Then
Label50.Visible = True
Shape1.Visible = True
Else
Label50.Visible = False
Shape1.Visible = False
End If

End Sub

Private Sub UkurA()

If DataGrid2.Columns(17) <> 0 Then
Text4.Text = DataGrid2.Columns(17)
Text5.Text = DataGrid2.Columns(18)
Text6.Text = DataGrid2.Columns(21)
Text7.Text = DataGrid2.Columns(22)
End If
End Sub

Private Sub UkurB()

If DataGrid3.Columns(25) <> 0 Then
Text9.Text = DataGrid3.Columns(25)
Text10.Text = DataGrid3.Columns(26)
Text14.Text = DataGrid3.Columns(29)
Text15.Text = DataGrid3.Columns(30)
End If
End Sub

Private Sub TampilUkur()
If DataGrid2.Columns(2) = Text1 Then
Call UkurA
End If

If DataGrid3.Columns(2) = Text1 Then
Call UkurB
End If

End Sub

Private Sub SimpanData()
KriteriaHitung

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

If Combo4 = "" Then
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
rs2.Open "select*from LKO", con
Dim SQLTambah As String
SQLTambah = "insert into LKO (Tahun,Bulan,Tanggal,NoMesin,Shift,NIK,Sequen,UrutanKanban,Waktu,AwalPengerjaan,AkhirPengerjaan,CutL,LotIDWire,Kombinasi,TermA,SealA,TermB,SealB,LotIDTermA,FCHTermA,FCWTermA,RCHTermA,RCWTermA,LotIDTermB,FCHTermB,FCWTermB,RCHTermB,RCWTermB,KodeDefect,QtyDefectOperator,QtyDefectMesin,QtyProduct,No4M,NoIssue,NoLogin,UrutanMesin) values ('" _
& Label37 & "','" & Label38 & "','" & Text11 & "','" & Combo1 & "','" & Combo2 & "','" & Combo4 & "','" & Text1 & "','" & Text28 & "','" & Label30 & "','" & Text25 & "','" & Text26 & "','" & Text20 & "','" & Text2 & "','" & Label25 & "','" & Text18 & "','" & Label5 & "','" & Text19 & "','" & Label6 & "','" & Text3 & "','" & Text4 & "','" & Text6 & "','" & Text5 & "','" & Text7 & "','" & Text8 & "','" & Text9 & "','" & Text14 & "','" & Text10 & "','" & Text15 & "','" & Combo3 & "','" & Text16 & "','" & Text27 & "','" & Text17 & "','" & Text21 & "','" & Text24 & "','" & Label41 & "','" & Text35 & "')"

con.Execute SQLTambah

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Jumlah
End Sub

Private Sub BukaUlang()

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus

Unload CrimpStandart
Unload CrimpStandartFull

Unload Me

WaktuTunggu

FrmOP.Show

Text28.Text = ""

Text35.Text = DataGrid4.Columns(35)
Text1 = DataGrid1.Columns(0)
Text17 = DataGrid1.Columns(4)
Text25.Text = DataGrid1.Columns(2)
Text26.Text = DataGrid1.Columns(3)
Text27.Text = DataGrid1.Columns(5)
Text28.Text = DataGrid1.Columns(1)

Combo2.Text = DataGrid4.Columns(4)
Combo4.Text = DataGrid4.Columns(5)

Jumlah

Label48 = total1
Label28 = total
Label31 = TeA
Label32 = TeB

Label49.Visible = False

End Sub

Public Sub AngkaLengkap()

Dim Tambah As Integer
Tambah = Len(Text1.Text)
If Tambah = 1 Then
Label51 = "000" + Text1.Text
Else
If Tambah = 2 Then
Label51 = "00" + Text1.Text
Else
If Tambah = 3 Then
Label51 = "0" + Text1.Text
Else
If Tambah = 4 Then
Label51 = Text1.Text
Else
If Tambah = 5 Then
Label51 = Text1.Text
Else
Label51 = Text1.Text
End If
End If
End If
End If
End If

End Sub

Private Sub Timer4_Timer()

If rs1 Is Nothing Then
Set rs1 = New ADODB.Recordset
End If

If rs1.State = adStateClosed Then
rs1.Open "select*from Prdlog", kon, adOpenDynamic, adLockOptimistic
End If

rs1.MoveNext

End Sub

Private Sub Otomatis()

If Text1.Text = Text36.Text Then
Text4.Text = DataGrid2.Columns(17)
Text5.Text = DataGrid2.Columns(18)
Text6.Text = DataGrid2.Columns(21)
Text7.Text = DataGrid2.Columns(22)

ElseIf Text1.Text <> Text36.Text Then
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = "0"
Text7.Text = "0"

End If

If Text1.Text = Text37.Text Then
Text9.Text = DataGrid3.Columns(25)
Text10.Text = DataGrid3.Columns(26)
Text14.Text = DataGrid3.Columns(29)
Text15.Text = DataGrid3.Columns(30)

ElseIf Text1.Text <> Text37.Text Then
Text9.Text = "0"
Text10.Text = "0"
Text14.Text = "0"
Text15.Text = "0"
End If


End Sub
