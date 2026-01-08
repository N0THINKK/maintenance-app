VERSION 5.00
Begin VB.Form DailyReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Report "
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      Caption         =   "Buka"
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
      Left            =   5040
      TabIndex        =   222
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Text142 
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
      Left            =   9600
      TabIndex        =   220
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "TOTAL"
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
      TabIndex        =   219
      Top             =   6120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   120
   End
   Begin VB.CommandButton Command14 
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
      Left            =   9960
      TabIndex        =   218
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Update"
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
      Left            =   7680
      TabIndex        =   217
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command12 
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
      Height          =   495
      Left            =   2400
      TabIndex        =   216
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command11 
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
      Height          =   495
      Left            =   120
      TabIndex        =   215
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Text141 
      Height          =   285
      Left            =   2880
      TabIndex        =   213
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text140 
      Height          =   285
      Left            =   2880
      TabIndex        =   212
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text139 
      Height          =   285
      Left            =   2880
      TabIndex        =   211
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text138 
      Height          =   285
      Left            =   2880
      TabIndex        =   210
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text137 
      Height          =   285
      Left            =   2880
      TabIndex        =   209
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text136 
      Height          =   285
      Left            =   2880
      TabIndex        =   208
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text135 
      Height          =   285
      Left            =   2880
      TabIndex        =   207
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text134 
      Height          =   285
      Left            =   2880
      TabIndex        =   206
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text133 
      Height          =   285
      Left            =   2880
      TabIndex        =   205
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text132 
      Height          =   285
      Left            =   2880
      TabIndex        =   204
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text131 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2880
      TabIndex        =   203
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text130 
      Height          =   285
      Left            =   9720
      TabIndex        =   201
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text129 
      Height          =   285
      Left            =   9720
      TabIndex        =   199
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   645
      Left            =   6720
      TabIndex        =   197
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text128 
      Height          =   645
      Left            =   4440
      TabIndex        =   196
      Top             =   6720
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   480
      TabIndex        =   194
      Top             =   2640
      Width           =   255
   End
   Begin VB.OptionButton Option10 
      Height          =   255
      Left            =   480
      TabIndex        =   193
      Top             =   5880
      Width           =   255
   End
   Begin VB.OptionButton Option9 
      Height          =   255
      Left            =   480
      TabIndex        =   192
      Top             =   5520
      Width           =   255
   End
   Begin VB.OptionButton Option8 
      Height          =   255
      Left            =   480
      TabIndex        =   191
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton Option7 
      Height          =   255
      Left            =   480
      TabIndex        =   190
      Top             =   4800
      Width           =   255
   End
   Begin VB.OptionButton Option6 
      Height          =   255
      Left            =   480
      TabIndex        =   189
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton Option5 
      Height          =   255
      Left            =   480
      TabIndex        =   188
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Height          =   255
      Left            =   480
      TabIndex        =   187
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Height          =   255
      Left            =   480
      TabIndex        =   186
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   480
      TabIndex        =   185
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   360
      TabIndex        =   184
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "-"
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
      Left            =   7800
      TabIndex        =   183
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "+"
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
      Left            =   7440
      TabIndex        =   182
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
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
      Left            =   7080
      TabIndex        =   181
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
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
      Left            =   6720
      TabIndex        =   180
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
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
      Left            =   6360
      TabIndex        =   179
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
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
      Left            =   6000
      TabIndex        =   178
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
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
      Left            =   5520
      TabIndex        =   177
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
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
      TabIndex        =   176
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
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
      Left            =   4800
      TabIndex        =   175
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
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
      Left            =   4440
      TabIndex        =   174
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text127 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   10440
      TabIndex        =   173
      Top             =   6120
      Width           =   550
   End
   Begin VB.TextBox Text126 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9720
      TabIndex        =   172
      Top             =   6120
      Width           =   550
   End
   Begin VB.TextBox Text125 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9000
      TabIndex        =   171
      Top             =   6120
      Width           =   550
   End
   Begin VB.TextBox Text124 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   8280
      TabIndex        =   170
      Top             =   6120
      Width           =   550
   End
   Begin VB.TextBox Text123 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   7440
      TabIndex        =   169
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text122 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6720
      TabIndex        =   168
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text121 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6000
      TabIndex        =   167
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text120 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5160
      TabIndex        =   166
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text119 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4440
      TabIndex        =   165
      Top             =   6120
      Width           =   675
   End
   Begin VB.TextBox Text118 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3720
      TabIndex        =   164
      Top             =   6120
      Width           =   580
   End
   Begin VB.TextBox Text117 
      Height          =   285
      Left            =   10440
      TabIndex        =   163
      Top             =   5760
      Width           =   550
   End
   Begin VB.TextBox Text116 
      Height          =   285
      Left            =   10440
      TabIndex        =   162
      Top             =   5400
      Width           =   550
   End
   Begin VB.TextBox Text115 
      Height          =   285
      Left            =   10440
      TabIndex        =   161
      Top             =   5040
      Width           =   550
   End
   Begin VB.TextBox Text114 
      Height          =   285
      Left            =   10440
      TabIndex        =   160
      Top             =   4680
      Width           =   550
   End
   Begin VB.TextBox Text113 
      Height          =   285
      Left            =   10440
      TabIndex        =   159
      Top             =   4320
      Width           =   550
   End
   Begin VB.TextBox Text112 
      Height          =   285
      Left            =   10440
      TabIndex        =   158
      Top             =   3960
      Width           =   550
   End
   Begin VB.TextBox Text111 
      Height          =   285
      Left            =   10440
      TabIndex        =   157
      Top             =   3600
      Width           =   550
   End
   Begin VB.TextBox Text110 
      Height          =   285
      Left            =   10440
      TabIndex        =   156
      Top             =   3240
      Width           =   550
   End
   Begin VB.TextBox Text109 
      Height          =   285
      Left            =   10440
      TabIndex        =   155
      Top             =   2880
      Width           =   550
   End
   Begin VB.TextBox Text108 
      Height          =   285
      Left            =   10440
      TabIndex        =   154
      Top             =   2520
      Width           =   550
   End
   Begin VB.TextBox Text107 
      Height          =   285
      Left            =   9720
      TabIndex        =   153
      Top             =   5760
      Width           =   550
   End
   Begin VB.TextBox Text106 
      Height          =   285
      Left            =   9720
      TabIndex        =   152
      Top             =   5400
      Width           =   550
   End
   Begin VB.TextBox Text105 
      Height          =   285
      Left            =   9720
      TabIndex        =   151
      Top             =   5040
      Width           =   550
   End
   Begin VB.TextBox Text104 
      Height          =   285
      Left            =   9720
      TabIndex        =   150
      Top             =   4680
      Width           =   550
   End
   Begin VB.TextBox Text103 
      Height          =   285
      Left            =   9720
      TabIndex        =   149
      Top             =   4320
      Width           =   550
   End
   Begin VB.TextBox Text102 
      Height          =   285
      Left            =   9720
      TabIndex        =   148
      Top             =   3960
      Width           =   550
   End
   Begin VB.TextBox Text101 
      Height          =   285
      Left            =   9720
      TabIndex        =   147
      Top             =   3600
      Width           =   550
   End
   Begin VB.TextBox Text100 
      Height          =   285
      Left            =   9720
      TabIndex        =   146
      Top             =   3240
      Width           =   550
   End
   Begin VB.TextBox Text99 
      Height          =   285
      Left            =   9720
      TabIndex        =   145
      Top             =   2880
      Width           =   550
   End
   Begin VB.TextBox Text98 
      Height          =   285
      Left            =   9720
      TabIndex        =   144
      Top             =   2520
      Width           =   550
   End
   Begin VB.TextBox Text97 
      Height          =   285
      Left            =   9000
      TabIndex        =   143
      Top             =   5760
      Width           =   550
   End
   Begin VB.TextBox Text96 
      Height          =   285
      Left            =   9000
      TabIndex        =   142
      Top             =   5400
      Width           =   550
   End
   Begin VB.TextBox Text95 
      Height          =   285
      Left            =   9000
      TabIndex        =   141
      Top             =   5040
      Width           =   550
   End
   Begin VB.TextBox Text94 
      Height          =   285
      Left            =   9000
      TabIndex        =   140
      Top             =   4680
      Width           =   550
   End
   Begin VB.TextBox Text93 
      Height          =   285
      Left            =   9000
      TabIndex        =   139
      Top             =   4320
      Width           =   550
   End
   Begin VB.TextBox Text92 
      Height          =   285
      Left            =   9000
      TabIndex        =   138
      Top             =   3960
      Width           =   550
   End
   Begin VB.TextBox Text91 
      Height          =   285
      Left            =   9000
      TabIndex        =   137
      Top             =   3600
      Width           =   550
   End
   Begin VB.TextBox Text90 
      Height          =   285
      Left            =   9000
      TabIndex        =   136
      Top             =   3240
      Width           =   550
   End
   Begin VB.TextBox Text89 
      Height          =   285
      Left            =   9000
      TabIndex        =   135
      Top             =   2880
      Width           =   550
   End
   Begin VB.TextBox Text88 
      Height          =   285
      Left            =   9000
      TabIndex        =   134
      Top             =   2520
      Width           =   550
   End
   Begin VB.TextBox Text87 
      Height          =   285
      Left            =   8280
      TabIndex        =   133
      Top             =   5760
      Width           =   550
   End
   Begin VB.TextBox Text86 
      Height          =   285
      Left            =   8280
      TabIndex        =   132
      Top             =   5400
      Width           =   550
   End
   Begin VB.TextBox Text85 
      Height          =   285
      Left            =   8280
      TabIndex        =   131
      Top             =   5040
      Width           =   550
   End
   Begin VB.TextBox Text84 
      Height          =   285
      Left            =   8280
      TabIndex        =   130
      Top             =   4680
      Width           =   550
   End
   Begin VB.TextBox Text83 
      Height          =   285
      Left            =   8280
      TabIndex        =   129
      Top             =   4320
      Width           =   550
   End
   Begin VB.TextBox Text82 
      Height          =   285
      Left            =   8280
      TabIndex        =   128
      Top             =   3960
      Width           =   550
   End
   Begin VB.TextBox Text81 
      Height          =   285
      Left            =   8280
      TabIndex        =   127
      Top             =   3600
      Width           =   550
   End
   Begin VB.TextBox Text80 
      Height          =   285
      Left            =   8280
      TabIndex        =   126
      Top             =   3240
      Width           =   550
   End
   Begin VB.TextBox Text79 
      Height          =   285
      Left            =   8280
      TabIndex        =   125
      Top             =   2880
      Width           =   550
   End
   Begin VB.TextBox Text78 
      Height          =   285
      Left            =   8280
      TabIndex        =   124
      Top             =   2520
      Width           =   550
   End
   Begin VB.TextBox Text77 
      Height          =   285
      Left            =   7440
      TabIndex        =   118
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text76 
      Height          =   285
      Left            =   7440
      TabIndex        =   117
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text75 
      Height          =   285
      Left            =   7440
      TabIndex        =   116
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text74 
      Height          =   285
      Left            =   7440
      TabIndex        =   115
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text73 
      Height          =   285
      Left            =   7440
      TabIndex        =   114
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text72 
      Height          =   285
      Left            =   7440
      TabIndex        =   113
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text71 
      Height          =   285
      Left            =   7440
      TabIndex        =   112
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text70 
      Height          =   285
      Left            =   7440
      TabIndex        =   111
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text69 
      Height          =   285
      Left            =   7440
      TabIndex        =   110
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text68 
      Height          =   285
      Left            =   7440
      TabIndex        =   109
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text67 
      Height          =   285
      Left            =   6720
      TabIndex        =   108
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text66 
      Height          =   285
      Left            =   6720
      TabIndex        =   107
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text65 
      Height          =   285
      Left            =   6720
      TabIndex        =   106
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text64 
      Height          =   285
      Left            =   6720
      TabIndex        =   105
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text63 
      Height          =   285
      Left            =   6720
      TabIndex        =   104
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text62 
      Height          =   285
      Left            =   6720
      TabIndex        =   103
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text61 
      Height          =   285
      Left            =   6720
      TabIndex        =   102
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text60 
      Height          =   285
      Left            =   6720
      TabIndex        =   101
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text59 
      Height          =   285
      Left            =   6720
      TabIndex        =   100
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text58 
      Height          =   285
      Left            =   6720
      TabIndex        =   99
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text57 
      Height          =   285
      Left            =   6000
      TabIndex        =   98
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text56 
      Height          =   285
      Left            =   6000
      TabIndex        =   97
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text55 
      Height          =   285
      Left            =   6000
      TabIndex        =   96
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text54 
      Height          =   285
      Left            =   6000
      TabIndex        =   95
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text53 
      Height          =   285
      Left            =   6000
      TabIndex        =   94
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text52 
      Height          =   285
      Left            =   6000
      TabIndex        =   93
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text51 
      Height          =   285
      Left            =   6000
      TabIndex        =   92
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text50 
      Height          =   285
      Left            =   6000
      TabIndex        =   91
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text49 
      Height          =   285
      Left            =   6000
      TabIndex        =   90
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text48 
      Height          =   285
      Left            =   6000
      TabIndex        =   89
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text47 
      Height          =   285
      Left            =   5160
      TabIndex        =   88
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text46 
      Height          =   285
      Left            =   5160
      TabIndex        =   87
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text45 
      Height          =   285
      Left            =   5160
      TabIndex        =   86
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text44 
      Height          =   285
      Left            =   5160
      TabIndex        =   85
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text43 
      Height          =   285
      Left            =   5160
      TabIndex        =   84
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text42 
      Height          =   285
      Left            =   5160
      TabIndex        =   83
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text41 
      Height          =   285
      Left            =   5160
      TabIndex        =   82
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text40 
      Height          =   285
      Left            =   5160
      TabIndex        =   81
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text39 
      Height          =   285
      Left            =   5160
      TabIndex        =   80
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text38 
      Height          =   285
      Left            =   5160
      TabIndex        =   79
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text37 
      Height          =   285
      Left            =   4440
      TabIndex        =   78
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox Text36 
      Height          =   285
      Left            =   4440
      TabIndex        =   77
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox Text34 
      Height          =   285
      Left            =   4440
      TabIndex        =   76
      Top             =   5040
      Width           =   675
   End
   Begin VB.TextBox Text33 
      Height          =   285
      Left            =   4440
      TabIndex        =   75
      Top             =   4680
      Width           =   675
   End
   Begin VB.TextBox Text32 
      Height          =   285
      Left            =   4440
      TabIndex        =   74
      Top             =   4320
      Width           =   675
   End
   Begin VB.TextBox Text31 
      Height          =   285
      Left            =   4440
      TabIndex        =   73
      Top             =   3960
      Width           =   675
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   4440
      TabIndex        =   72
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   4440
      TabIndex        =   71
      Top             =   3240
      Width           =   675
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   4440
      TabIndex        =   70
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   4440
      TabIndex        =   69
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   3720
      TabIndex        =   60
      Top             =   5760
      Width           =   580
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   3720
      TabIndex        =   59
      Top             =   5400
      Width           =   580
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   3720
      TabIndex        =   58
      Top             =   5040
      Width           =   580
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   3720
      TabIndex        =   57
      Top             =   4680
      Width           =   580
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   3720
      TabIndex        =   56
      Top             =   4320
      Width           =   580
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   3720
      TabIndex        =   55
      Top             =   3960
      Width           =   580
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   3720
      TabIndex        =   54
      Top             =   3600
      Width           =   580
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   3720
      TabIndex        =   53
      Top             =   3240
      Width           =   580
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3720
      TabIndex        =   52
      Top             =   2880
      Width           =   580
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   3720
      TabIndex        =   51
      Top             =   2520
      Width           =   580
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   2160
      TabIndex        =   49
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   405
      Left            =   2160
      TabIndex        =   48
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   2160
      TabIndex        =   47
      Top             =   6120
      Width           =   580
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2160
      TabIndex        =   45
      Top             =   5760
      Width           =   580
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2160
      TabIndex        =   44
      Top             =   5400
      Width           =   580
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2160
      TabIndex        =   43
      Top             =   5040
      Width           =   580
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2160
      TabIndex        =   42
      Top             =   4680
      Width           =   580
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2160
      TabIndex        =   41
      Top             =   4320
      Width           =   580
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2160
      TabIndex        =   40
      Top             =   3960
      Width           =   580
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2160
      TabIndex        =   39
      Top             =   3600
      Width           =   580
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   38
      Top             =   3240
      Width           =   580
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   37
      Top             =   2880
      Width           =   580
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   36
      Top             =   2520
      Width           =   580
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
      Height          =   285
      Left            =   8040
      TabIndex        =   18
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
      TabIndex        =   4
      Top             =   840
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
      ItemData        =   "DailyReport.frx":0000
      Left            =   1680
      List            =   "DailyReport.frx":001F
      TabIndex        =   3
      Top             =   840
      Width           =   1575
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
      ItemData        =   "DailyReport.frx":0076
      Left            =   4680
      List            =   "DailyReport.frx":0083
      TabIndex        =   2
      Top             =   840
      Width           =   1335
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
      Left            =   6240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
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
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Jam Kerja"
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
      Left            =   9600
      TabIndex        =   221
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY CUTTING (PCS)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   214
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      Caption         =   "Waktu Monitor"
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
      TabIndex        =   202
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      Caption         =   "Waktu Mulai"
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
      TabIndex        =   200
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Efficiency Ratio"
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
      Left            =   6600
      TabIndex        =   198
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      Caption         =   "TARGET / JAM"
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
      TabIndex        =   195
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   OTHER"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   123
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tunggu Kanban"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   122
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mesin Troble"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   121
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tunggu Material"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   120
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NON SET-UP PROSES"
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
      TabIndex        =   119
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sisi B"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   68
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   67
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sisi A"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   66
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APPL/TERM"
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
      Left            =   5880
      TabIndex        =   65
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIDE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   64
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COLOR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   63
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WIRE"
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
      Left            =   4440
      TabIndex        =   62
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SET-UP PROSES"
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
      Left            =   4440
      TabIndex        =   61
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Support MH (Menit)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   50
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Over Time"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   46
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Excluding Time"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   35
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "GL"
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
      TabIndex        =   34
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Cutting QTY/JAM"
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
      Left            =   960
      TabIndex        =   33
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "14:30 - 15:25/15:40"
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
      Left            =   960
      TabIndex        =   32
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "11:30/12:30 - 13:30"
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
      Left            =   960
      TabIndex        =   31
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "18:15 - 19:00"
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
      Left            =   960
      TabIndex        =   30
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "17:00 - 18:00"
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
      Left            =   960
      TabIndex        =   29
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "15:40 - 16:40"
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
      Left            =   960
      TabIndex        =   28
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "13:30 -14:30"
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
      Left            =   960
      TabIndex        =   27
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "10:45 - 11:30/12:30"
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
      Left            =   960
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "09:30 - 10:45"
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
      Left            =   960
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "08:30 - 09:30"
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
      Left            =   960
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "07:30 - 08:30"
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
      Left            =   960
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Standart Jam Kerja"
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
      Left            =   720
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "N O R M A  L                                       T   I  M E"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jam Kerja"
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
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
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
      TabIndex        =   19
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Daily Report Proses Cutting"
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
      Left            =   4320
      TabIndex        =   17
      Top             =   120
      Width           =   3615
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
      TabIndex        =   16
      Top             =   600
      Width           =   1095
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
      Left            =   6240
      TabIndex        =   15
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
      Left            =   1920
      TabIndex        =   14
      Top             =   600
      Width           =   1215
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
      Left            =   4680
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label30 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label38 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   10440
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label37 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   120
      Width           =   375
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
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label43 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label42 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label39 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label40 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      Height          =   735
      Left            =   3360
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "DailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset

Dim qty1
Dim qty2
Dim qty3
Dim qty4
Dim qty5
Dim qty6
Dim qty7
Dim qty8
Dim qty9
Dim qty10
Dim qty11

Dim side1 As Long
Dim side2
Dim side3
Dim side4
Dim side5
Dim side6
Dim side7
Dim side8
Dim side9
Dim side10
Dim side11

Dim color1
Dim color2
Dim color3
Dim color4
Dim color5
Dim color6
Dim color7
Dim color8
Dim color9
Dim color10
Dim color11

Dim A1
Dim A2
Dim A3
Dim A4
Dim A5
Dim A6
Dim A7
Dim A8
Dim A9
Dim A10
Dim A11

Dim B1
Dim B2
Dim B3
Dim B4
Dim B5
Dim B6
Dim B7
Dim B8
Dim B9
Dim B10
Dim B11

Dim D1
Dim D2
Dim D3
Dim D4
Dim D5
Dim D6
Dim D7
Dim D8
Dim D9
Dim D10
Dim D11

Dim material1
Dim material2
Dim material3
Dim material4
Dim material5
Dim material6
Dim material7
Dim material8
Dim material9
Dim material10
Dim material11

Dim troble1
Dim troble2
Dim troble3
Dim troble4
Dim troble5
Dim troble6
Dim troble7
Dim troble8
Dim troble9
Dim troble10
Dim troble11

Dim kanban1
Dim kanban2
Dim kanban3
Dim kanban4
Dim kanban5
Dim kanban6
Dim kanban7
Dim kanban8
Dim kanban9
Dim kanban10
Dim kanban11

Dim other1
Dim other2
Dim other3
Dim other4
Dim other5
Dim other6
Dim other7
Dim other8
Dim other9
Dim other10
Dim other11

Private Sub Command1_Click()

If Option1 = True Then
side1 = side1 + 1
Text27.Text = side1
End If

If Option2 = True Then
side2 = side2 + 1
Text28.Text = side2
End If

If Option3 = True Then
side3 = side3 + 1
Text29.Text = side3
End If

If Option4 = True Then
side4 = side4 + 1
Text30.Text = side4
End If

If Option5 = True Then
side5 = side5 + 1
Text31.Text = side5
End If

If Option6 = True Then
side6 = side6 + 1
Text32.Text = side6
End If

If Option7 = True Then
side7 = side7 + 1
Text33.Text = side7
End If

If Option8 = True Then
side8 = side8 + 1
Text34.Text = side8
End If

If Option9 = True Then
side9 = side9 + 1
Text36.Text = side9
End If

If Option10 = True Then
side10 = side10 + 1
Text37.Text = side10
End If

End Sub

Private Sub Command10_Click()
If Option1 = True Then
D1 = D1 - 1
Text68.Text = D1
End If

If Option2 = True Then
D2 = D2 - 1
Text69.Text = D2
End If

If Option3 = True Then
D3 = D3 - 1
Text70.Text = D3
End If

If Option4 = True Then
D4 = D4 - 1
Text71.Text = D4
End If

If Option5 = True Then
D5 = D5 - 1
Text72.Text = D5
End If

If Option6 = True Then
D6 = D6 - 1
Text73.Text = D6
End If

If Option7 = True Then
D7 = D7 - 1
Text74.Text = D7
End If

If Option8 = True Then
D8 = D8 - 1
Text75.Text = D8
End If

If Option9 = True Then
D9 = D9 - 1
Text76.Text = D9
End If

If Option10 = True Then
D10 = D10 - 1
Text77.Text = D10
End If
End Sub

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command12_Click()
DataDailyReport.Show
End Sub

Private Sub Command13_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "select*from Dy", con, adOpenDynamic, adLockOptimistic
If Label37 = DataDailyReport.DataGrid1.Columns(0) And Label38 = DataDailyReport.DataGrid1.Columns(1) And Text11 = DataDailyReport.DataGrid1.Columns(2) And Combo1 = DataDailyReport.DataGrid1.Columns(4) And Combo2 = DataDailyReport.DataGrid1.Columns(6) Then
DataDailyReport.DataGrid1.Columns(10) = Text14
DataDailyReport.DataGrid1.Columns(11) = Text128
DataDailyReport.DataGrid1.Columns(12) = Text15
DataDailyReport.DataGrid1.Columns(13) = Text129
DataDailyReport.DataGrid1.Columns(14) = Text130
DataDailyReport.DataGrid1.Columns(15) = Text13
DataDailyReport.DataGrid1.Columns(16) = Text131
DataDailyReport.DataGrid1.Columns(17) = Text118
DataDailyReport.DataGrid1.Columns(18) = Text119
DataDailyReport.DataGrid1.Columns(19) = Text120
DataDailyReport.DataGrid1.Columns(20) = Text121
DataDailyReport.DataGrid1.Columns(21) = Text122
DataDailyReport.DataGrid1.Columns(22) = Text123
DataDailyReport.DataGrid1.Columns(23) = Text124
DataDailyReport.DataGrid1.Columns(24) = Text125
DataDailyReport.DataGrid1.Columns(25) = Text126
DataDailyReport.DataGrid1.Columns(26) = Text127
DataDailyReport.DataGrid1.Columns(27) = Text16

   MsgBox "Data sudah diupdate", vbInformation, "Informasi"
Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
'rs1.Update

'Unload Me

'DataDailyReport.Show
End Sub

Private Sub Command14_Click()
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

If Combo4 = "" Then
MsgBox " NIK belum terisi"
Exit Sub
End If

If Text1 = "" Then
MsgBox " Nama belum dipilih"
Exit Sub
End If

If Text15 = "" Then
MsgBox " Ratio Produk belum terisi"
Exit Sub
End If

Set rs2 = New ADODB.Recordset
rs2.Open "select*from Dy", con, adOpenDynamic, adLockOptimistic
Dim SQLTambah As String
SQLTambah = "insert into Dy (Tahun,Bulan,Tanggal,Waktu,NoMesin,NoUrut,Shift,NIK,Nama,JamKerja,Qtyjam,Targetjam,EfisiensiRatio,WaktuMulai,WaktuMonitor,ExcludingTime,QtyCutting,SupportMH,Side,Color,SisiA,SisiB,Dobel,TungguMaterial,MesinTrouble,TungguKanban,Other,GL) values ('" _
& Label37 & "','" & Label38 & "','" & Text11 & "','" & Label30 & "','" & Combo1 & "','" & Text35 & "','" & Combo2 & "','" & Combo4 & "','" & Text1 & "','" & Text142 & "','" & Text14 & "','" & Text128 & "','" & Text15 & "','" & Text129 & "','" & Text130 & "','" & Text13 & "','" & Text131 & "','" & Text118 & "','" & Text119 & "','" & Text120 & "','" & Text121 & "','" & Text122 & "','" & Text123 & "','" & Text124 & "','" & Text125 & "','" & Text126 & "','" & Text127 & "','" & Text16 & "')"

con.Execute SQLTambah

'
  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
End Sub

Private Sub Command15_Click()
side11 = side1 + side2 + side3 + side4 + side5 + side6 + side7 + side8 + side9 + side10
Text119 = side11

color11 = color1 + color2 + color3 + color4 + color5 + color6 + color7 + color8 + color9 + color10
Text120 = color11

A11 = A1 + A2 + A3 + A4 + A5 + A6 + A7 + A8 + A9 + A10
Text121 = A11

B11 = B1 + B2 + B3 + B4 + B5 + B6 + B7 + B8 + B9 + B10
Text122 = B11

D11 = D1 + D2 + D3 + D4 + D5 + D6 + D7 + D8 + D9 + D10
Text123 = D11

Text119 = Val(Text27) + Val(Text28) + Val(Text29) + Val(Text30) + Val(Text31) + Val(Text32) + Val(Text33) + Val(Text34) + Val(Text36) + Val(Text37)

Text120 = Val(Text38) + Val(Text39) + Val(Text40) + Val(Text41) + Val(Text42) + Val(Text43) + Val(Text44) + Val(Text45) + Val(Text46) + Val(Text47)

Text121 = Val(Text48) + Val(Text49) + Val(Text50) + Val(Text51) + Val(Text52) + Val(Text53) + Val(Text54) + Val(Text55) + Val(Text56) + Val(Text57)

Text122 = Val(Text58) + Val(Text59) + Val(Text60) + Val(Text61) + Val(Text62) + Val(Text63) + Val(Text64) + Val(Text65) + Val(Text66) + Val(Text67)

Text123 = Val(Text68) + Val(Text69) + Val(Text70) + Val(Text71) + Val(Text72) + Val(Text73) + Val(Text74) + Val(Text75) + Val(Text76) + Val(Text77)

Text124 = Val(Text78) + Val(Text79) + Val(Text80) + Val(Text81) + Val(Text82) + Val(Text83) + Val(Text84) + Val(Text85) + Val(Text86) + Val(Text87)

Text125 = Val(Text88) + Val(Text89) + Val(Text90) + Val(Text91) + Val(Text92) + Val(Text93) + Val(Text94) + Val(Text95) + Val(Text96) + Val(Text97)

Text126 = Val(Text98) + Val(Text99) + Val(Text100) + Val(Text101) + Val(Text102) + Val(Text103) + Val(Text104) + Val(Text105) + Val(Text106) + Val(Text107)

Text127 = Val(Text108) + Val(Text109) + Val(Text110) + Val(Text111) + Val(Text112) + Val(Text113) + Val(Text114) + Val(Text115) + Val(Text116) + Val(Text117)

Text131 = Val(Text132) + Val(Text133) + Val(Text134) + Val(Text135) + Val(Text136) + Val(Text137) + Val(Text138) + Val(Text139) + Val(Text140) + Val(Text141)


End Sub

Private Sub Command16_Click()
rs3.CursorLocation = adUseClient
rs3.Open "select * from Dy where Tahun = '" & Label37 & "' and Bulan = '" & Label38 & "' and Tanggal = '" & Text11 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'", con
If Not rs3.EOF Then
    Text14 = rs3!Qtyjam
    Text128 = rs3!Targetjam
    Text15 = rs3!EfisiensiRatio
    Text129 = rs3!WaktuMulai
    Text130 = rs3!WaktuMonitor
    Text13 = rs3!ExcludingTime
    Text131 = rs3!QtyCutting
    Text118 = rs3!SupportMH
    Text119 = rs3!Side
    Text120 = rs3!Color
    Text121 = rs3!SisiA
    Text122 = rs3!SisiB
    Text123 = rs3!Dobel
    Text124 = rs3!TungguMaterial
    Text125 = rs3!MesinTrouble
    Text126 = rs3!TungguKanban
    Text127 = rs3!Other
    
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Ada Informasi!!!!!"
End If
End Sub

Private Sub Command2_Click()

If Option1 = True Then
side1 = side1 - 1
Text27.Text = side1
End If

If Option2 = True Then
side2 = side2 - 1
Text28.Text = side2
End If

If Option3 = True Then
side3 = side3 - 1
Text29.Text = side3
End If

If Option4 = True Then
side4 = side4 - 1
Text30.Text = side4
End If

If Option5 = True Then
side5 = side5 - 1
Text31.Text = side5
End If

If Option6 = True Then
side6 = side6 - 1
Text32.Text = side6
End If

If Option7 = True Then
side7 = side7 - 1
Text33.Text = side7
End If

If Option8 = True Then
side8 = side8 - 1
Text34.Text = side8
End If

If Option9 = True Then
side9 = side9 - 1
Text36.Text = side9
End If

If Option10 = True Then
side10 = side10 - 1
Text37.Text = side10
End If
End Sub

Private Sub Command3_Click()

If Option1 = True Then
color1 = color1 + 1
Text38.Text = color1
End If

If Option2 = True Then
color2 = color2 + 1
Text39.Text = color2
End If

If Option3 = True Then
color3 = color3 + 1
Text40.Text = color3
End If

If Option4 = True Then
color4 = color4 + 1
Text41.Text = color4
End If

If Option5 = True Then
color5 = color5 + 1
Text42.Text = color5
End If

If Option6 = True Then
color6 = color6 + 1
Text43.Text = color6
End If

If Option7 = True Then
color7 = color7 + 1
Text44.Text = color7
End If

If Option8 = True Then
color8 = color8 + 1
Text45.Text = color8
End If

If Option9 = True Then
color9 = color9 + 1
Text46.Text = color9
End If

If Option10 = True Then
color10 = color10 + 1
Text47.Text = color10
End If
End Sub

Private Sub Command4_Click()
If Option1 = True Then
color1 = color1 - 1
Text38.Text = color1
End If

If Option2 = True Then
color2 = color2 - 1
Text39.Text = color2
End If

If Option3 = True Then
color3 = color3 - 1
Text40.Text = color3
End If

If Option4 = True Then
color4 = color4 - 1
Text41.Text = color4
End If

If Option5 = True Then
color5 = color5 - 1
Text42.Text = color5
End If

If Option6 = True Then
color6 = color6 - 1
Text43.Text = color6
End If

If Option7 = True Then
color7 = color7 - 1
Text44.Text = color7
End If

If Option8 = True Then
color8 = color8 - 1
Text45.Text = color8
End If

If Option9 = True Then
color9 = color9 - 1
Text46.Text = color9
End If

If Option10 = True Then
color10 = color10 - 1
Text47.Text = color10
End If
End Sub

Private Sub Command5_Click()

If Option1 = True Then
A1 = A1 + 1
Text48.Text = A1
End If

If Option2 = True Then
A2 = A2 + 1
Text49.Text = A2
End If

If Option3 = True Then
A3 = A3 + 1
Text50.Text = A3
End If

If Option4 = True Then
A4 = A4 + 1
Text51.Text = A4
End If

If Option5 = True Then
A5 = A5 + 1
Text52.Text = A5
End If

If Option6 = True Then
A6 = A6 + 1
Text53.Text = A6
End If

If Option7 = True Then
A7 = A7 + 1
Text54.Text = A7
End If

If Option8 = True Then
A8 = A8 + 1
Text55.Text = A8
End If

If Option9 = True Then
A9 = A9 + 1
Text56.Text = A9
End If

If Option10 = True Then
A10 = A10 + 1
Text57.Text = A10
End If
End Sub

Private Sub Command6_Click()
If Option1 = True Then
A1 = A1 - 1
Text48.Text = A1
End If

If Option2 = True Then
A2 = A2 - 1
Text49.Text = A2
End If

If Option3 = True Then
A3 = A3 - 1
Text50.Text = A3
End If

If Option4 = True Then
A4 = A4 - 1
Text51.Text = A4
End If

If Option5 = True Then
A5 = A5 - 1
Text52.Text = A5
End If

If Option6 = True Then
A6 = A6 - 1
Text53.Text = A6
End If

If Option7 = True Then
A7 = A7 - 1
Text54.Text = A7
End If

If Option8 = True Then
A8 = A8 - 1
Text55.Text = A8
End If

If Option9 = True Then
A9 = A9 - 1
Text56.Text = A9
End If

If Option10 = True Then
A10 = A10 - 1
Text57.Text = A10
End If
End Sub

Private Sub Command7_Click()

If Option1 = True Then
B1 = B1 + 1
Text58.Text = B1
End If

If Option2 = True Then
B2 = B2 + 1
Text59.Text = B2
End If

If Option3 = True Then
B3 = B3 + 1
Text60.Text = B3
End If

If Option4 = True Then
B4 = B4 + 1
Text61.Text = B4
End If

If Option5 = True Then
B5 = B5 + 1
Text62.Text = B5
End If

If Option6 = True Then
B6 = B6 + 1
Text63.Text = B6
End If

If Option7 = True Then
B7 = B7 + 1
Text64.Text = B7
End If

If Option8 = True Then
B8 = B8 + 1
Text65.Text = B8
End If

If Option9 = True Then
B9 = B9 + 1
Text66.Text = B9
End If

If Option10 = True Then
B10 = B10 + 1
Text67.Text = B10
End If
End Sub

Private Sub Command8_Click()
If Option1 = True Then
B1 = B1 - 1
Text58.Text = B1
End If

If Option2 = True Then
B2 = B2 - 1
Text59.Text = B2
End If

If Option3 = True Then
B3 = B3 - 1
Text60.Text = B3
End If

If Option4 = True Then
B4 = B4 - 1
Text61.Text = B4
End If

If Option5 = True Then
B5 = B5 - 1
Text62.Text = B5
End If

If Option6 = True Then
B6 = B6 - 1
Text63.Text = B6
End If

If Option7 = True Then
B7 = B7 - 1
Text64.Text = B7
End If

If Option8 = True Then
B8 = B8 - 1
Text65.Text = B8
End If

If Option9 = True Then
B9 = B9 - 1
Text66.Text = B9
End If

If Option10 = True Then
B10 = B10 - 1
Text67.Text = B10
End If
End Sub

Private Sub Command9_Click()
If Option1 = True Then
D1 = D1 + 1
Text68.Text = D1
End If

If Option2 = True Then
D2 = D2 + 1
Text69.Text = D2
End If

If Option3 = True Then
D3 = D3 + 1
Text70.Text = D3
End If

If Option4 = True Then
D4 = D4 + 1
Text71.Text = D4
End If

If Option5 = True Then
D5 = D5 + 1
Text72.Text = D5
End If

If Option6 = True Then
D6 = D6 + 1
Text73.Text = D6
End If

If Option7 = True Then
D7 = D7 + 1
Text74.Text = D7
End If

If Option8 = True Then
D8 = D8 + 1
Text75.Text = D8
End If

If Option9 = True Then
D9 = D9 + 1
Text76.Text = D9
End If

If Option10 = True Then
D10 = D10 + 1
Text77.Text = D10
End If
End Sub

Private Sub Form_Activate()



Label37 = Format(Now, "yyyy")
Label38 = Format(Now, "mm")


End Sub

Private Sub Form_Load()

JumlahAwal
Text11 = Format(Now, "yyyy/mm/dd")
Combo1.Text = NoMesin.Text21.Text
Text35.Text = Right$(Combo1.Text, 2)

'With Combo2
'.AddItem "A"
'.AddItem "B"
'.AddItem "NS"
'End With

tampil
Option1 = True

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\DataDailyReportCutting.mdb"
Set rs2 = New ADODB.Recordset
rs2.Open "select*from Dy", con, adOpenDynamic, adLockOptimistic


End Sub

Private Sub PilihJam()

End Sub

Private Sub JumlahAwal()
side1 = 0
side2 = 0
side3 = 0
side4 = 0
side5 = 0
side6 = 0
side7 = 0
side8 = 0
side9 = 0
side10 = 0
side11 = 0

color1 = 0
color2 = 0
color3 = 0
color4 = 0
color5 = 0
color6 = 0
color7 = 0
color8 = 0
color9 = 0
color10 = 0
color11 = 0

A1 = 0
A2 = 0
A3 = 0
A4 = 0
A5 = 0
A6 = 0
A7 = 0
A8 = 0
A9 = 0
A10 = 0
A11 = 0

B1 = 0
B2 = 0
B3 = 0
B4 = 0
B5 = 0
B6 = 0
B7 = 0
B8 = 0
B9 = 0
B10 = 0
B11 = 0

D1 = 0
D2 = 0
D3 = 0
D4 = 0
D5 = 0
D6 = 0
D7 = 0
D8 = 0
D9 = 0
D10 = 0
D11 = 0

material1 = 0
material2 = 0
material3 = 0
material4 = 0
material5 = 0
material6 = 0
material7 = 0
material8 = 0
material9 = 0
material10 = 0
material11 = 0

troble1 = 0
troble2 = 0
troble3 = 0
troble4 = 0
troble5 = 0
troble6 = 0
troble7 = 0
troble8 = 0
troble9 = 0
troble10 = 0
troble11 = 0

kanban1 = 0
kanban2 = 0
kanban3 = 0
kanban4 = 0
kanban5 = 0
kanban6 = 0
kanban7 = 0
kanban8 = 0
kanban9 = 0
kanban10 = 0
kanban11 = 0

other1 = 0
other2 = 0
other3 = 0
other4 = 0
other5 = 0
other6 = 0
other7 = 0
other8 = 0
other9 = 0
other10 = 0
other11 = 0

End Sub

Private Sub tampil()

Text27 = side1
Text28 = side2
Text29 = side3
Text30 = side4
Text31 = side5
Text32 = side6
Text33 = side7
Text34 = side8
Text36 = side9
Text37 = side10
Text119 = side11

Text38 = color1
Text39 = color2
Text40 = color3
Text41 = color4
Text42 = color5
Text43 = color6
Text44 = color7
Text45 = color8
Text46 = color9
Text47 = color10
Text120 = color11

Text48 = A1
Text49 = A2
Text50 = A3
Text51 = A4
Text52 = A5
Text53 = A6
Text54 = A7
Text55 = A8
Text56 = A9
Text57 = A10
Text121 = A11

Text58 = B1
Text59 = B2
Text60 = B3
Text61 = B4
Text62 = B5
Text63 = B6
Text64 = B7
Text65 = B8
Text66 = B9
Text67 = B10
Text122 = B11

Text68 = D1
Text69 = D2
Text70 = D3
Text71 = D4
Text72 = D5
Text73 = D6
Text74 = D7
Text75 = D8
Text76 = D9
Text77 = D10
Text123 = D11

Text78 = material1
Text79 = material2
Text80 = material3
Text81 = material4
Text82 = material5
Text83 = material6
Text84 = material7
Text85 = material8
Text86 = material9
Text87 = material10
Text124 = material11

Text88 = troble1
Text89 = troble2
Text90 = troble3
Text91 = troble4
Text92 = troble5
Text93 = troble6
Text94 = troble7
Text95 = troble8
Text96 = troble9
Text97 = troble10
Text125 = troble11

Text98 = kanban1
Text99 = kanban2
Text100 = kanban3
Text101 = kanban4
Text102 = kanban5
Text103 = kanban6
Text104 = kanban7
Text105 = kanban8
Text106 = kanban9
Text107 = kanban10
Text126 = kanban11

Text108 = other1
Text109 = other2
Text110 = other3
Text111 = other4
Text112 = other5
Text113 = other6
Text114 = other7
Text115 = other8
Text116 = other9
Text117 = other10
Text127 = other11

Text131 = 0
Text132 = 0
Text133 = 0
Text134 = 0
Text135 = 0
Text136 = 0
Text137 = 0
Text138 = 0
Text139 = 0
Text140 = 0
Text141 = 0

End Sub

Private Sub Timer1_Timer()
Label30 = Format(Now, "hh:mm:ss")
End Sub

Private Sub Jumlah()

End Sub
