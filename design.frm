VERSION 5.00
Begin VB.Form Design 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Circuit"
   ClientHeight    =   4050
   ClientLeft      =   1530
   ClientTop       =   2400
   ClientWidth     =   6825
   Icon            =   "design.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4050
   ScaleWidth      =   6825
   Visible         =   0   'False
   Begin VB.TextBox helpText 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   206
      Text            =   "design.frx":000C
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   199
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   205
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   198
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   204
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   197
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   203
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   196
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   202
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   195
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   201
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   194
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   200
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   193
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   199
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   192
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   198
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   191
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   197
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   190
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   196
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   189
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   195
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   188
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   194
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   187
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   193
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   186
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   192
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   185
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   191
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   184
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   190
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   183
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   189
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   182
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   188
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   181
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   187
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   180
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   186
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   179
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   185
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   178
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   184
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   177
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   183
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   176
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   182
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   175
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   181
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   174
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   180
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   173
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   179
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   172
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   178
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   171
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   177
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   170
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   176
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   169
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   175
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   168
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   174
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   167
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   173
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   166
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   172
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   165
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   171
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   164
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   170
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   163
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   169
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   162
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   168
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   161
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   167
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   160
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   166
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   159
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   165
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   158
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   164
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   157
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   163
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   156
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   162
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   155
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   161
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   154
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   160
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   153
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   159
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   152
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   158
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   151
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   157
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   150
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   156
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   149
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   155
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   148
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   154
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   147
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   153
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   146
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   152
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   145
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   151
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   144
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   150
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   143
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   149
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   142
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   148
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   141
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   147
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   140
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   146
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   139
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   145
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   138
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   144
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   137
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   143
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   136
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   142
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   135
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   141
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   134
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   140
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   133
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   139
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   132
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   138
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   131
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   137
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   130
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   136
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   129
      Left            =   6912
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   135
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   128
      Left            =   6552
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   134
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   127
      Left            =   6192
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   133
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   126
      Left            =   5832
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   132
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   125
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   131
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   124
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   130
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   123
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   129
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   122
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   128
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   121
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   127
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   120
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   126
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   119
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   125
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   118
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   124
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   117
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   123
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   116
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   122
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   115
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   121
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   114
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   120
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   113
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   119
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   112
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   118
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   111
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   117
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   110
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   116
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   109
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   115
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   108
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   114
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   107
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   113
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   106
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   112
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   105
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   111
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   104
      Left            =   5112
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   110
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   103
      Left            =   4752
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   109
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   102
      Left            =   4392
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   108
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   101
      Left            =   4032
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   107
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   100
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   106
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   99
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   105
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   98
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   104
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   97
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   103
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   96
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   102
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   95
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   101
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   94
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   100
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   93
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   99
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   92
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   98
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   91
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   97
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   90
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   96
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   89
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   95
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   88
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   94
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   87
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   93
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   86
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   92
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   85
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   91
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   84
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   90
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   83
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   89
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   82
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   88
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   81
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   87
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   80
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   86
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   79
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   85
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   78
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   84
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   77
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   83
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   76
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   82
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   75
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   81
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   74
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   80
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   73
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   79
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   72
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   78
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   71
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   77
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   70
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   76
      Top             =   1920
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   69
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   75
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   68
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   74
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   67
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   73
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   66
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   72
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   65
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   71
      Top             =   3360
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   64
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   70
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   63
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   69
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   62
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   68
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   61
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   67
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   60
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   66
      Top             =   3000
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   59
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   65
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   58
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   64
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   57
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   63
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   56
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   62
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   55
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   61
      Top             =   2640
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   54
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   60
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   53
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   59
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   52
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   58
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   51
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   57
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   50
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   56
      Top             =   2280
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   49
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   55
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   48
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   54
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   47
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   53
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   46
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   52
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   45
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   51
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   44
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   50
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   43
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   49
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   42
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   48
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   41
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   47
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   40
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   46
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   39
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   45
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   38
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   44
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   37
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   43
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   36
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   42
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   35
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   41
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   34
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   40
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   33
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   39
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   32
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   38
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   31
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   37
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   30
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   36
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   9
      Left            =   3312
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   35
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   8
      Left            =   2952
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   34
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   7
      Left            =   2592
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   33
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   6
      Left            =   2232
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   32
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   5
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   31
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   4
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   30
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   29
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   28
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   27
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   26
      Top             =   120
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   29
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   25
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   28
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   24
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   27
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   23
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   26
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   22
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   25
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   21
      Top             =   1560
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   24
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   20
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   23
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   19
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   22
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   18
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   21
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   17
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   20
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   16
      Top             =   1200
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   19
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   15
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   18
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   14
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   17
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   13
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   16
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   15
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   11
      Top             =   840
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   14
      Left            =   1512
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   10
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   13
      Left            =   1152
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   9
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   12
      Left            =   792
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   8
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   11
      Left            =   432
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   7
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Tile 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   10
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   480
      Width           =   396
   End
   Begin VB.PictureBox Border 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   468
      Index           =   3
      Left            =   132
      ScaleHeight     =   465
      ScaleWidth      =   1605
      TabIndex        =   5
      Top             =   4584
      Width           =   1608
   End
   Begin VB.PictureBox Border 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   408
      Index           =   2
      Left            =   912
      ScaleHeight     =   405
      ScaleWidth      =   5580
      TabIndex        =   4
      Top             =   4512
      Width           =   5580
   End
   Begin VB.PictureBox Border 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3588
      Index           =   1
      Left            =   4056
      ScaleHeight     =   3585
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   4608
      Width           =   180
   End
   Begin VB.PictureBox Border 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1536
      Index           =   0
      Left            =   6372
      ScaleHeight     =   1530
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   4500
      Width           =   180
   End
   Begin VB.ListBox datalist 
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1296
   End
   Begin VB.ListBox values 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "Design"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Dragged As Boolean
Public dx, dy As Single
Public T0L, T0T As Single
Public bwith, cvs, changed
Public DesignID
Public SelectedIndex
Public ActiveObject
Private helpTextCount
Public Sub init()
bwith = 100
sze = CELL_SPACING
With Me
.Width = sze * 5 + 100
.Height = sze * 5 + 300
.Caption = "Preparing Design Area.."
.values.Clear
.datalist.Clear
.Border(0).Width = 20 * sze
.Border(1).Width = 20 * sze
.Border(2).Width = bwith
.Border(3).Width = bwith
.Border(0).Height = bwith
.Border(1).Height = bwith
.Border(2).Height = 10 * sze
.Border(3).Height = 10 * sze
.Border(0).Top = 0
.Border(1).Top = 10 * sze + bwith
.Border(2).Top = 0 'left
.Border(3).Top = 0 'right
.Border(0).Left = 0 'top
.Border(1).Left = 0 'bottom
.Border(2).Left = 0
.Border(3).Left = 20 * sze + bwith
End With
For i = 0 To 199
 'If I Then Load Tile(I)
 If i Mod 10 = 0 Then DoEvents
 With Tile(i)
  .Visible = True
  .Left = sze * (i Mod 20) + bwith
  .Top = sze * (i \ 20) + bwith
  .Width = sze
  .Height = sze
  .BorderStyle = 1
  .Enabled = True
 End With
  values.AddItem "0"
  datalist.AddItem "-1"
Next
Caption = "Mesh.cct"
If IsEmpty(DesignID) Then DesignID = -1
Show
End Sub
Public Sub drop(X, Y, Shift)
On Error GoTo ThrowException
X = X - Me.Left
Y = Y - Me.Top
If X > Me.Width Or Y > Me.Height Then Exit Sub
For i = 0 To 199
 With Tile(i)
  test = X >= .Left And X < .Left + .Width
  test = test And Y >= .Top And Y < .Top + .Height
 End With
 If test Then Exit For
Next
If i = 200 Then Exit Sub

  If Not canvas(DD.Cform).ActiveObject Is Nothing Then
   Dim components(200)
   Dim properties(200)
   For i = 0 To canvas(DD.Cform).datalist.ListCount - 1
    components(i) = canvas(DD.Cform).datalist.list(i)
    properties(i) = canvas(DD.Cform).values.list(i)
   Next
   canvas(DD.Cform).ActiveObject.Insert components, properties, i, Toolbox.finalsel
  End If

X = i Mod 20: Y = i \ 20
ReDim ARR(20, 10) As Integer
For i = 0 To datalist.ListCount - 1
ARR(i Mod 20, i \ 20) = Val(datalist.list(i))
Next
changed = True
MDI.Refresh
cc = ARR(X, Y)
If cc = -1 Or Shift Then
    'pref2.Hide
    typ = Toolbox.finalsel
    Select Case typ
    Case 2, 6, 8, 5, 1
        If ARR(X + 1, Y) = 22 Then Beep: typ = -1
    Case 0, 2, 8, 3, 1
        If ARR(X, Y + 1) = 22 Then Beep: typ = -1
    Case Is < 10: op = "0"
    Case 10, 11: op = "R"
    Case 12, 13: op = "-CJ"
    Case 14, 15: op = "LJ"
    Case 16, 17: op = "V,0"
    Case 18 To 21: op = ".7,0+0.025/I"
    Case 22
     If X > 1 And Y > 1 And ARR(X - 1, Y) <> 22 And ARR(X, Y - 1) <> 22 Then
      op = "50K,0+0.025/I,Hfe"
      ARR(X - 1, Y) = 4
      ARR(X, Y - 1) = 9
      Else: Beep: typ = -1
     End If
    End Select
    values.list(X + Y * 20) = op
    ARR(X, Y) = typ
    datalist.Clear
    For Y = 0 To 9
    For X = 0 To 19
    datalist.AddItem Str(ARR(X, Y))
    Next X, Y
Else
   If cc < 10 Then
     datalist.list(X + Y * 20) = "-1"
   Else
    'If Val(pref2.Ctyp) = cc And cc <> 22 And Val(pref2.Ref) <> X + Y * 20 Then
    'If cc Mod 2 = 0 Then cc = cc + 1 Else cc = cc - 1
    'End If
    'pref2.Ref = Str(X + Y * 20)
    'pref2.Ctyp = Str(cc)
    'pref2.ReNew
    'pref2.Show
   End If
End If
ReNew
Exit Sub
ThrowException:
 MsgBox Err.Description
End Sub
Public Sub ReNew()
 clearHelpTip
 For i = 0 To 199
  bol = Val(datalist.list(i))
  If bol = -1 Then
   bol = 99
  Else
   If bol > 9 And MDI.mnuLabels.Checked Then
    newHelpTip Tile(i), values.list(i)
   End If
  End If
  Tile(i).Picture = Toolbox.Sym(bol).Picture
 Next
End Sub
Public Sub drag(X, Y)
If Abs(X - dx) + Abs(Y - dy) > 200 Then dx = X: dy = Y
nx = Tile(0).Left + X - dx
ny = Tile(0).Top + Y - dy
test = nx <= bwith * -1 And ny <= bwith * -1
test = test And nx > Tile(0).Width * -15
test = test And ny > Tile(0).Height * -5
If test Then X = dx: Y = dy
For i = 0 To 3
 With Border(i)
  .Visible = False
  .Left = .Left + X - dx
  .Top = .Top + Y - dy
 End With
Next
' speedup patch only move tile if within design(a)
For i = 0 To 199
 With Tile(i)
  If .Left < Me.Width And .Top < Me.Height Then
   .Left = .Left + X - dx
   .Top = .Top + Y - dy
  End If
 End With
Next
For i = 0 To 3
 With Border(i)
  .Visible = True
   End With
Next
End Sub

Public Sub tmup(Button As Integer, Index As Integer, X As Single, Y As Single, Shift As Integer)
On Error GoTo ThrowException
Select Case Button
Case 1
    X = X + Tile(Index).Left
    Y = Y + Tile(Index).Top
    If X < Me.Width And Y < Me.Height Then
     For i = 0 To 199
      With Tile(i)
       If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height And i <> Index Then
        test = True: Exit For
       End If
      End With
     Next
    End If
    If Not ActiveObject Is Nothing Then
     Dim components(200)
     Dim properties(200)
     For i = 0 To datalist.ListCount - 1
      components(i) = datalist.list(i)
      properties(i) = values.list(i)
     Next
     If Not ActiveObject.Delete(components, properties, Index) Then
      test = False
      MsgBox "This circuit does not support this action"
     End If
     If Not ActiveObject.Insert(components, properties, i, components(Index)) Then
      test = False
      MsgBox "This circuit does not support this action"
     End If
    End If
    
    If test Then
     Tile(i).Picture = Tile(Index).Picture
     Tile(Index).Picture = Toolbox.Sym(99).Picture
     datalist.list(i) = datalist.list(Index)
     values.list(i) = values.list(Index)
     datalist.list(Index) = "-1"
     changed = True
     MDI.Refresh
     ReNew
    End If
    sze = CELL_SPACING
    With Tile(Index)
     .Left = sze * (Index Mod 20) + T0L
     .Top = sze * (Index \ 20) + T0T
     .Width = sze
     .Height = sze
    End With
End Select
Exit Sub
ThrowException:
 MsgBox Err.Description
     sze = CELL_SPACING
    With Tile(Index)
     .Left = sze * (Index Mod 20) + T0L
     .Top = sze * (Index \ 20) + T0T
     .Width = sze
     .Height = sze
    End With
End Sub
Public Sub keymove(Index As Integer, KeyCode As Integer)
Dim dx, dy As Integer
sze = CELL_SPACING
Select Case KeyCode
 Case 37: If Tile(0).Left > bwith Then dx = -1
 Case 38: If Tile(0).Top > bwith Then dy = -1
 Case 39: If Tile(19).Left < sze * 20 Then dx = 1
 Case 40: If Tile(19).Top < sze * 20 Then dy = 1
 Case 36 'home
  For z = 0 To 199
   With Tile(z)
    .Left = (z Mod 20) * .Width + bwith
    .Top = (z \ 20) * .Width + bwith
   End With
   For i = 0 To 3
    Border(i).Top = 0
    Border(i).Left = 0
   Next i
   Border(1).Top = 10 * sze
   Border(3).Left = 20 * sze
  Next
 Case 46 ' delete
   datalist.list(Index) = "-1"
   Tile(Index).Picture = Toolbox.Sym(99).Picture
End Select
If KeyCode > 36 And KeyCode < 41 Then
 For i = 1 To sze Step 100
  For z = 0 To 199
   With Tile(z)
    .Left = .Left + (dx) * 100
    .Top = .Top + (dy) * 100
   End With
  Next
  For z = 0 To 3
   With Border(z)
    .Left = .Left + dx * 100
    .Top = .Top + dy * 100
   End With
  Next
 Next
End If
End Sub




Private Sub Form_Activate()
DD.Cform = cvs
MDI.Refresh
End Sub

Private Sub Form_Load()
 Me.Visible = True
 DoEvents
 init
 Set ActiveObject = Nothing
End Sub

Private Sub tile_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 keymove Index, KeyCode
End Sub
Private Sub Tile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
dx = X
dy = Y
Dragged = False
T0L = Tile(0).Left
T0T = Tile(0).Top
 If Button = 2 Then
  MDI.updatePopupMenu datalist.list(Index)
  Me.SelectedIndex = Index
  Me.PopupMenu MDI.mnuPopup
  MDI.popupLeft = X + Tile(Index).Left
  MDI.popupTop = Y + Tile(Index).Top
 End If
 
End Sub
Private Sub Tile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
  MDI.updatePopupMenu datalist.list(Index)
  Me.SelectedIndex = Index
  Me.PopupMenu MDI.mnuPopup
Case 1
    With Tile(Index)
    .Tag = "dragged"
    .Left = .Left + X - dx
    .Top = .Top + Y - dy
    .ZOrder 0
    End With
End Select
End Sub
Private Sub Tile_mouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Tile(Index).Tag = "dragged" Then
  tmup Button, Index, X, Y, Shift
  Tile(Index).Tag = ""
 Else
  drop Tile(Index).Left, Tile(Index).Top, 0
 End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
 If Me.Caption = "Preparing Design Area.." Then
  Cancel = True
  Exit Sub
 End If
 If Me.Caption <> "closing..." Then MDI.mnuFileClose_Click
End Sub

Public Function newHelpTip(parentObj As Object, text)
 X = parentObj.Left + parentObj.Width * 0.75
 Y = parentObj.Top + parentObj.Height * 0.25
 helpTextCount = helpTextCount + 1
 Load helpText(helpTextCount)
 helpText(helpTextCount).Visible = True
 helpText(helpTextCount).text = text
 helpText(helpTextCount).Top = Y
 helpText(helpTextCount).Left = X
 helpText(helpTextCount).ZOrder 0
 helpText(helpTextCount).Width = Len(text) * helpText(helpTextCount).Font.Size * Screen.TwipsPerPixelX
End Function
Public Function clearHelpTip()
 For i = 1 To helpTextCount
  Unload helpText(i)
 Next
 helpTextCount = 0
End Function


