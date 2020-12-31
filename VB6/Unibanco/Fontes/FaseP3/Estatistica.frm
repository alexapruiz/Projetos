VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Estatistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acompanhamento da Produção"
   ClientHeight    =   8328
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   11556
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8328
   ScaleWidth      =   11556
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   3996
      TabIndex        =   3
      Top             =   7824
      Width           =   1512
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   288
      Top             =   5352
   End
   Begin VB.Frame Frame1 
      Height          =   528
      Left            =   228
      TabIndex        =   98
      Top             =   7080
      Width           =   5184
      Begin VB.OptionButton optFiltro 
         Caption         =   "Fininvest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3732
         TabIndex        =   144
         Top             =   192
         Width           =   1110
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Malotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2610
         TabIndex        =   2
         Top             =   192
         Width           =   1005
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   192
         Width           =   1230
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   192
         Width           =   948
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   390
      Left            =   5556
      TabIndex        =   4
      Top             =   7824
      Width           =   1512
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   5628
      Left            =   180
      OleObjectBlob   =   "Estatistica.frx":0000
      TabIndex        =   165
      Top             =   1044
      Width           =   5196
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   2
      Left            =   5628
      TabIndex        =   193
      Top             =   852
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   3
      Left            =   5628
      TabIndex        =   192
      Top             =   1092
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   4
      Left            =   5628
      TabIndex        =   191
      Top             =   1332
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   5
      Left            =   5628
      TabIndex        =   190
      Top             =   1572
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   6
      Left            =   5628
      TabIndex        =   189
      Top             =   1812
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   7
      Left            =   5628
      TabIndex        =   188
      Top             =   2052
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   8
      Left            =   5628
      TabIndex        =   187
      Top             =   2292
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   9
      Left            =   5628
      TabIndex        =   186
      Top             =   2532
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   10
      Left            =   5628
      TabIndex        =   185
      Top             =   2772
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   11
      Left            =   5628
      TabIndex        =   184
      Top             =   4692
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   12
      Left            =   5628
      TabIndex        =   183
      Top             =   4932
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   13
      Left            =   5628
      TabIndex        =   182
      Top             =   5652
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   14
      Left            =   5628
      TabIndex        =   181
      Top             =   5892
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   15
      Left            =   5628
      TabIndex        =   180
      Top             =   6132
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   16
      Left            =   5628
      TabIndex        =   179
      Top             =   6372
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   17
      Left            =   5628
      TabIndex        =   178
      Top             =   6612
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   18
      Left            =   5628
      TabIndex        =   177
      Top             =   6852
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   19
      Left            =   5628
      TabIndex        =   176
      Top             =   5172
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   20
      Left            =   5628
      TabIndex        =   175
      Top             =   5412
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   21
      Left            =   5628
      TabIndex        =   174
      Top             =   3252
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   22
      Left            =   5628
      TabIndex        =   173
      Top             =   3492
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   23
      Left            =   5628
      TabIndex        =   172
      Top             =   3732
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   24
      Left            =   5628
      TabIndex        =   171
      Top             =   3972
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   25
      Left            =   5628
      TabIndex        =   170
      Top             =   3012
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   26
      Left            =   5628
      TabIndex        =   169
      Top             =   4212
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   27
      Left            =   5628
      TabIndex        =   168
      Top             =   4452
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   1
      Left            =   5628
      TabIndex        =   167
      Top             =   612
      Width           =   408
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Height          =   230
      Index           =   0
      Left            =   5628
      TabIndex        =   166
      Top             =   370
      Width           =   408
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   0
      Left            =   9468
      TabIndex        =   164
      Top             =   372
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   0
      Left            =   10008
      TabIndex        =   163
      Top             =   372
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   0
      Left            =   10908
      TabIndex        =   162
      Top             =   372
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   0
      Left            =   8700
      TabIndex        =   161
      Top             =   372
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recepcionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   0
      Left            =   6060
      TabIndex        =   160
      Top             =   372
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   1
      Left            =   8700
      TabIndex        =   159
      Top             =   612
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   1
      Left            =   10908
      TabIndex        =   158
      Top             =   612
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   1
      Left            =   10008
      TabIndex        =   157
      Top             =   612
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   1
      Left            =   9468
      TabIndex        =   156
      Top             =   612
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Complementação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   1
      Left            =   6060
      TabIndex        =   155
      Top             =   612
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   27
      Left            =   8700
      TabIndex        =   154
      Top             =   4452
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   27
      Left            =   10908
      TabIndex        =   153
      Top             =   4452
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   27
      Left            =   10008
      TabIndex        =   152
      Top             =   4452
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   27
      Left            =   9468
      TabIndex        =   151
      Top             =   4452
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em CSP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   27
      Left            =   6060
      TabIndex        =   150
      Top             =   4452
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   26
      Left            =   8700
      TabIndex        =   149
      Top             =   4212
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   26
      Left            =   10908
      TabIndex        =   148
      Top             =   4212
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   26
      Left            =   10008
      TabIndex        =   147
      Top             =   4212
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   26
      Left            =   9468
      TabIndex        =   146
      Top             =   4212
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para CSP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   26
      Left            =   6060
      TabIndex        =   145
      Top             =   4212
      Width           =   2604
   End
   Begin VB.Label LblAgProc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2856"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   3240
      TabIndex        =   143
      Top             =   450
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data do Movimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   510
      TabIndex        =   142
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label lblDataProc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17/07/2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   3240
      TabIndex        =   141
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agência Processadora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   510
      TabIndex        =   140
      Top             =   450
      Width           =   2700
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   25
      Left            =   8700
      TabIndex        =   139
      Top             =   3012
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   25
      Left            =   10908
      TabIndex        =   138
      Top             =   3012
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   25
      Left            =   10008
      TabIndex        =   137
      Top             =   3012
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   25
      Left            =   9468
      TabIndex        =   136
      Top             =   3012
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Estorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   25
      Left            =   6060
      TabIndex        =   135
      Top             =   3012
      Width           =   2604
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Correção de Ag. Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   24
      Left            =   6060
      TabIndex        =   134
      Top             =   3972
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   24
      Left            =   9468
      TabIndex        =   133
      Top             =   3972
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   24
      Left            =   10008
      TabIndex        =   132
      Top             =   3972
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   24
      Left            =   10908
      TabIndex        =   131
      Top             =   3972
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   24
      Left            =   8700
      TabIndex        =   130
      Top             =   3972
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Correção de Ag. Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   23
      Left            =   6060
      TabIndex        =   129
      Top             =   3732
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   23
      Left            =   9468
      TabIndex        =   128
      Top             =   3732
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   23
      Left            =   10008
      TabIndex        =   127
      Top             =   3732
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   23
      Left            =   10908
      TabIndex        =   126
      Top             =   3732
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   23
      Left            =   8700
      TabIndex        =   125
      Top             =   3732
      Width           =   732
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   22
      Left            =   8700
      TabIndex        =   124
      Top             =   3492
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   22
      Left            =   10908
      TabIndex        =   123
      Top             =   3492
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   22
      Left            =   10008
      TabIndex        =   122
      Top             =   3492
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   22
      Left            =   9468
      TabIndex        =   121
      Top             =   3492
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Confirmação Ag/Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   22
      Left            =   6060
      TabIndex        =   120
      Top             =   3492
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   21
      Left            =   8700
      TabIndex        =   119
      Top             =   3252
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   21
      Left            =   10908
      TabIndex        =   118
      Top             =   3252
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   21
      Left            =   10008
      TabIndex        =   117
      Top             =   3252
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   21
      Left            =   9468
      TabIndex        =   116
      Top             =   3252
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Confirmação Ag/Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   21
      Left            =   6060
      TabIndex        =   115
      Top             =   3252
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   20
      Left            =   8700
      TabIndex        =   114
      Top             =   5412
      Width           =   732
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   20
      Left            =   10908
      TabIndex        =   113
      Top             =   5412
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   20
      Left            =   10008
      TabIndex        =   112
      Top             =   5412
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   20
      Left            =   9468
      TabIndex        =   111
      Top             =   5412
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Recaptura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   20
      Left            =   6060
      TabIndex        =   110
      Top             =   5412
      Width           =   2604
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Recaptura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   19
      Left            =   6060
      TabIndex        =   109
      Top             =   5172
      Width           =   2604
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   19
      Left            =   8700
      TabIndex        =   108
      Top             =   5172
      Width           =   732
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   19
      Left            =   9468
      TabIndex        =   107
      Top             =   5172
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   19
      Left            =   10008
      TabIndex        =   106
      Top             =   5172
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   19
      Left            =   10908
      TabIndex        =   105
      Top             =   5172
      Width           =   492
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   18
      Left            =   10908
      TabIndex        =   7
      Top             =   6852
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   18
      Left            =   10008
      TabIndex        =   104
      Top             =   6852
      Width           =   876
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   18
      Left            =   9468
      TabIndex        =   103
      Top             =   6852
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   18
      Left            =   8700
      TabIndex        =   102
      Top             =   6852
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Excluído / Ocorrência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   18
      Left            =   6060
      TabIndex        =   101
      Top             =   6852
      Width           =   2604
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:30:35"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1140
      TabIndex        =   100
      Top             =   7836
      Width           =   720
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17/07/2000"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   156
      TabIndex        =   99
      Top             =   7824
      Width           =   936
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   6060
      TabIndex        =   97
      Top             =   7200
      Width           =   2616
   End
   Begin VB.Label lblTotPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   10908
      TabIndex        =   96
      Top             =   7200
      Width           =   492
   End
   Begin VB.Label lblTotDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   10008
      TabIndex        =   95
      Top             =   7200
      Width           =   876
   End
   Begin VB.Label lblTotPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   9468
      TabIndex        =   94
      Top             =   7200
      Width           =   492
   End
   Begin VB.Label lblTotCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   8700
      TabIndex        =   93
      Top             =   7200
      Width           =   732
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   10908
      TabIndex        =   92
      Top             =   132
      Width           =   492
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Doctos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   10008
      TabIndex        =   91
      Top             =   132
      Width           =   876
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   9468
      TabIndex        =   90
      Top             =   132
      Width           =   492
   End
   Begin VB.Label lblFiltro 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   8700
      TabIndex        =   89
      Top             =   132
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   6060
      TabIndex        =   88
      Top             =   132
      Width           =   2604
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   17
      Left            =   6060
      TabIndex        =   87
      Top             =   6612
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   17
      Left            =   9468
      TabIndex        =   86
      Top             =   6612
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   17
      Left            =   10008
      TabIndex        =   85
      Top             =   6612
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   17
      Left            =   10908
      TabIndex        =   84
      Top             =   6612
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   17
      Left            =   8700
      TabIndex        =   83
      Top             =   6612
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Expedição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   16
      Left            =   6060
      TabIndex        =   82
      Top             =   6372
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   16
      Left            =   9468
      TabIndex        =   81
      Top             =   6372
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   16
      Left            =   10008
      TabIndex        =   80
      Top             =   6372
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   16
      Left            =   10908
      TabIndex        =   79
      Top             =   6372
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   16
      Left            =   8700
      TabIndex        =   78
      Top             =   6372
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Expedição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   15
      Left            =   6060
      TabIndex        =   77
      Top             =   6132
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   15
      Left            =   9468
      TabIndex        =   76
      Top             =   6132
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   15
      Left            =   10008
      TabIndex        =   75
      Top             =   6132
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   15
      Left            =   10908
      TabIndex        =   74
      Top             =   6132
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   15
      Left            =   8700
      TabIndex        =   73
      Top             =   6132
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Transmissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   14
      Left            =   6060
      TabIndex        =   72
      Top             =   5892
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   14
      Left            =   9468
      TabIndex        =   71
      Top             =   5892
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   14
      Left            =   10008
      TabIndex        =   70
      Top             =   5892
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   14
      Left            =   10908
      TabIndex        =   69
      Top             =   5892
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   14
      Left            =   8700
      TabIndex        =   68
      Top             =   5892
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Transmissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   13
      Left            =   6060
      TabIndex        =   67
      Top             =   5652
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   13
      Left            =   9468
      TabIndex        =   66
      Top             =   5652
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   13
      Left            =   10008
      TabIndex        =   65
      Top             =   5652
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   13
      Left            =   10908
      TabIndex        =   64
      Top             =   5652
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   13
      Left            =   8700
      TabIndex        =   63
      Top             =   5652
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Alçada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   12
      Left            =   6060
      TabIndex        =   62
      Top             =   4932
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   12
      Left            =   9468
      TabIndex        =   61
      Top             =   4932
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   12
      Left            =   10008
      TabIndex        =   60
      Top             =   4932
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   12
      Left            =   10908
      TabIndex        =   59
      Top             =   4932
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   12
      Left            =   8700
      TabIndex        =   58
      Top             =   4932
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Alçada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   11
      Left            =   6060
      TabIndex        =   57
      Top             =   4692
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   11
      Left            =   9468
      TabIndex        =   56
      Top             =   4692
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   11
      Left            =   10008
      TabIndex        =   55
      Top             =   4692
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   11
      Left            =   10908
      TabIndex        =   54
      Top             =   4692
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   11
      Left            =   8700
      TabIndex        =   53
      Top             =   4692
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Troca de Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   10
      Left            =   6060
      TabIndex        =   52
      Top             =   2772
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   10
      Left            =   9468
      TabIndex        =   51
      Top             =   2772
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   10
      Left            =   10008
      TabIndex        =   50
      Top             =   2772
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   10
      Left            =   10908
      TabIndex        =   49
      Top             =   2772
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   10
      Left            =   8700
      TabIndex        =   48
      Top             =   2772
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Ilegíveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   9
      Left            =   6060
      TabIndex        =   47
      Top             =   2532
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   9
      Left            =   9468
      TabIndex        =   46
      Top             =   2532
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   9
      Left            =   10008
      TabIndex        =   45
      Top             =   2532
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   9
      Left            =   10908
      TabIndex        =   44
      Top             =   2532
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   9
      Left            =   8700
      TabIndex        =   43
      Top             =   2532
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Ilegíveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   8
      Left            =   6060
      TabIndex        =   42
      Top             =   2292
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   8
      Left            =   9468
      TabIndex        =   41
      Top             =   2292
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   8
      Left            =   10008
      TabIndex        =   40
      Top             =   2292
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   8
      Left            =   10908
      TabIndex        =   39
      Top             =   2292
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   8
      Left            =   8700
      TabIndex        =   38
      Top             =   2292
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Prova Zero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   7
      Left            =   6060
      TabIndex        =   37
      Top             =   2052
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   7
      Left            =   9468
      TabIndex        =   36
      Top             =   2052
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   7
      Left            =   10008
      TabIndex        =   35
      Top             =   2052
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   7
      Left            =   10908
      TabIndex        =   34
      Top             =   2052
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   7
      Left            =   8700
      TabIndex        =   33
      Top             =   2052
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Prova Zero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   6
      Left            =   6060
      TabIndex        =   32
      Top             =   1812
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   6
      Left            =   9468
      TabIndex        =   31
      Top             =   1812
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   6
      Left            =   10008
      TabIndex        =   30
      Top             =   1812
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   6
      Left            =   10908
      TabIndex        =   29
      Top             =   1812
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   6
      Left            =   8700
      TabIndex        =   28
      Top             =   1812
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Vínculo Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   5
      Left            =   6060
      TabIndex        =   27
      Top             =   1572
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   5
      Left            =   9468
      TabIndex        =   26
      Top             =   1572
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   5
      Left            =   10008
      TabIndex        =   25
      Top             =   1572
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   5
      Left            =   10908
      TabIndex        =   24
      Top             =   1572
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   5
      Left            =   8700
      TabIndex        =   23
      Top             =   1572
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Vínculo Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   4
      Left            =   6060
      TabIndex        =   22
      Top             =   1332
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   4
      Left            =   9468
      TabIndex        =   21
      Top             =   1332
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   4
      Left            =   10008
      TabIndex        =   20
      Top             =   1332
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   4
      Left            =   10908
      TabIndex        =   19
      Top             =   1332
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   4
      Left            =   8700
      TabIndex        =   18
      Top             =   1332
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Vínc. Automático"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   3
      Left            =   6060
      TabIndex        =   17
      Top             =   1092
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   3
      Left            =   9468
      TabIndex        =   16
      Top             =   1092
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   3
      Left            =   10008
      TabIndex        =   15
      Top             =   1092
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   3
      Left            =   10908
      TabIndex        =   14
      Top             =   1092
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   3
      Left            =   8700
      TabIndex        =   13
      Top             =   1092
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Complementação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   2
      Left            =   6060
      TabIndex        =   12
      Top             =   852
      Width           =   2604
   End
   Begin VB.Label lblPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   2
      Left            =   9468
      TabIndex        =   11
      Top             =   852
      Width           =   492
   End
   Begin VB.Label lblQtdDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.234.567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   2
      Left            =   10008
      TabIndex        =   10
      Top             =   852
      Width           =   876
   End
   Begin VB.Label lblPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   2
      Left            =   10908
      TabIndex        =   9
      Top             =   852
      Width           =   492
   End
   Begin VB.Label lblQtdCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123.456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Index           =   2
      Left            =   8700
      TabIndex        =   8
      Top             =   852
      Width           =   732
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   5556
      TabIndex        =   6
      Top             =   24
      Width           =   5928
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   120
      TabIndex        =   5
      Top             =   24
      Width           =   5412
   End
End
Attribute VB_Name = "Estatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Matriz(1 To 28, 1 To 2) As Long

Private qryGetEstatistica As rdoQuery
Private qryGetTotalCapa As rdoQuery
Private qryGetTotalDocumento As rdoQuery
Private rsEstatistica As rdoResultset
Private rsTotalCapa As rdoResultset
Private rsTotalDoc As rdoResultset
Private Function FormataQuantidade(ByVal Qtd As Long) As String
    Dim strValor As String
    Dim strResult As String
    Dim Count As Integer
    
    strValor = Trim(str(Qtd))
    
    For Count = 1 To Len(strValor)
        strResult = Mid(strValor, Len(strValor) - Count + 1, 1) & strResult
        If (Count Mod 3 = 0) And (Count < Len(strValor)) Then
            strResult = "." & strResult
        End If
    Next
    If Len(strResult) = 0 Then
        strResult = "0"
    End If
    FormataQuantidade = strResult
End Function
Private Function ObtemTotalCapa(ByVal IdEnv_Mal As String) As Long
    On Error GoTo ErroTotalCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetTotalCapa.rdoParameters(0) = Geral.DataProcessamento
    qryGetTotalCapa.rdoParameters(1) = IdEnv_Mal
    Set rsTotalCapa = qryGetTotalCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsTotalCapa.EOF Then
        ObtemTotalCapa = 0
    Else
        ObtemTotalCapa = rsTotalCapa!Total
    End If
    rsTotalCapa.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroTotalCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do total de Envelopes/Malotes.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemTotalCapa = -1

End Function
Private Function ObtemTotalDocumento(ByVal IdEnv_Mal As String) As Long
    On Error GoTo ErroTotalDoc
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetTotalDocumento.rdoParameters(0) = Geral.DataProcessamento
    qryGetTotalDocumento.rdoParameters(1) = IdEnv_Mal
    Set rsTotalDoc = qryGetTotalDocumento.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsTotalDoc.EOF Then
        ObtemTotalDocumento = 0
    Else
        ObtemTotalDocumento = rsTotalDoc!Total
    End If
    rsTotalDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroTotalDoc:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do total de Documentos.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemTotalDocumento = -1

End Function
Private Function ObtemEstatistica(ByVal IdEnv_Mal As String) As Boolean
    On Error GoTo ErroEstat
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetEstatistica.rdoParameters(0) = Geral.DataProcessamento
    qryGetEstatistica.rdoParameters(1) = IdEnv_Mal
    Set rsEstatistica = qryGetEstatistica.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    ObtemEstatistica = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroEstat:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da Estatística de Envelopes/Malotes.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemEstatistica = False

End Function
Private Sub Zera_Matriz()
    Dim Count As Integer
    
    For Count = 1 To 28
        Matriz(Count, 1) = 0
        Matriz(Count, 2) = 0
    Next
    
End Sub
Private Sub Preenche_Matriz()
    
    Zera_Matriz
    
    If rsEstatistica.RowCount > 0 Then
        rsEstatistica.MoveFirst
    End If
    
    While Not rsEstatistica.EOF
        Select Case rsEstatistica!Status
            Case "0" 'Recepcionado
                Matriz(1, 1) = Matriz(1, 1) + rsEstatistica!QtdCapa
                Matriz(1, 2) = Matriz(1, 2) + rsEstatistica!QtdDoc
            Case "1" 'Para Complementacao
                Matriz(2, 1) = Matriz(2, 1) + rsEstatistica!QtdCapa
                Matriz(2, 2) = Matriz(2, 2) + rsEstatistica!QtdDoc
            Case "2" 'Em Complementacao
                Matriz(3, 1) = Matriz(3, 1) + rsEstatistica!QtdCapa
                Matriz(3, 2) = Matriz(3, 2) + rsEstatistica!QtdDoc
            Case "8", "9" 'Para Vinculo Automatico
                Matriz(4, 1) = Matriz(4, 1) + rsEstatistica!QtdCapa
                Matriz(4, 2) = Matriz(4, 2) + rsEstatistica!QtdDoc
            Case "7" 'Para Vinculo Manual
                Matriz(5, 1) = Matriz(5, 1) + rsEstatistica!QtdCapa
                Matriz(5, 2) = Matriz(5, 2) + rsEstatistica!QtdDoc
            Case "J" 'Em Vinculo Manual
                Matriz(6, 1) = Matriz(6, 1) + rsEstatistica!QtdCapa
                Matriz(6, 2) = Matriz(6, 2) + rsEstatistica!QtdDoc
            Case "4" 'Para Prova Zero
                Matriz(7, 1) = Matriz(7, 1) + rsEstatistica!QtdCapa
                Matriz(7, 2) = Matriz(7, 2) + rsEstatistica!QtdDoc
            Case "G" 'Em Prova Zero
                Matriz(8, 1) = Matriz(8, 1) + rsEstatistica!QtdCapa
                Matriz(8, 2) = Matriz(8, 2) + rsEstatistica!QtdDoc
            Case "5" 'Para Ilegiveis
                Matriz(9, 1) = Matriz(9, 1) + rsEstatistica!QtdCapa
                Matriz(9, 2) = Matriz(9, 2) + rsEstatistica!QtdDoc
            Case "H" 'Em Ilegivies
                Matriz(10, 1) = Matriz(10, 1) + rsEstatistica!QtdCapa
                Matriz(10, 2) = Matriz(10, 2) + rsEstatistica!QtdDoc
            Case "O" 'Em Troca de Ordem
                Matriz(11, 1) = Matriz(11, 1) + rsEstatistica!QtdCapa
                Matriz(11, 2) = Matriz(11, 2) + rsEstatistica!QtdDoc
            Case "6" 'Para Alcada
                Matriz(12, 1) = Matriz(12, 1) + rsEstatistica!QtdCapa
                Matriz(12, 2) = Matriz(12, 2) + rsEstatistica!QtdDoc
            Case "I" 'Em Alcada
                Matriz(13, 1) = Matriz(13, 1) + rsEstatistica!QtdCapa
                Matriz(13, 2) = Matriz(13, 2) + rsEstatistica!QtdDoc
            Case "P", "R" 'Para Transmissao
                Matriz(14, 1) = Matriz(14, 1) + rsEstatistica!QtdCapa
                Matriz(14, 2) = Matriz(14, 2) + rsEstatistica!QtdDoc
            Case "S" 'Em Transmissao
                Matriz(15, 1) = Matriz(15, 1) + rsEstatistica!QtdCapa
                Matriz(15, 2) = Matriz(15, 2) + rsEstatistica!QtdDoc
            Case "T" 'Para Expedicao
                Matriz(16, 1) = Matriz(16, 1) + rsEstatistica!QtdCapa
                Matriz(16, 2) = Matriz(16, 2) + rsEstatistica!QtdDoc
            Case "K" 'Em Expedicao
                Matriz(17, 1) = Matriz(17, 1) + rsEstatistica!QtdCapa
                Matriz(17, 2) = Matriz(17, 2) + rsEstatistica!QtdDoc
            Case "E" 'Expedido
                Matriz(18, 1) = Matriz(18, 1) + rsEstatistica!QtdCapa
                Matriz(18, 2) = Matriz(18, 2) + rsEstatistica!QtdDoc
            Case "D", "F", "X" 'Excluido / Ocorrencia
                Matriz(19, 1) = Matriz(19, 1) + rsEstatistica!QtdCapa
                Matriz(19, 2) = Matriz(19, 2) + rsEstatistica!QtdDoc
            Case "A" 'Para Recaptura
                Matriz(20, 1) = Matriz(20, 1) + rsEstatistica!QtdCapa
                Matriz(20, 2) = Matriz(20, 2) + rsEstatistica!QtdDoc
            Case "B" 'Em Recaptura
                Matriz(21, 1) = Matriz(21, 1) + rsEstatistica!QtdCapa
                Matriz(21, 2) = Matriz(21, 2) + rsEstatistica!QtdDoc
            Case "L" 'Para Confirmação de Agência / Conta
                Matriz(22, 1) = Matriz(22, 1) + rsEstatistica!QtdCapa
                Matriz(22, 2) = Matriz(22, 2) + rsEstatistica!QtdDoc
            Case "M" 'Em   Confirmação de Agência / Conta
                Matriz(23, 1) = Matriz(23, 1) + rsEstatistica!QtdCapa
                Matriz(23, 2) = Matriz(23, 2) + rsEstatistica!QtdDoc
            Case "Y" 'Para Correção de Ag. Conta
                Matriz(24, 1) = Matriz(24, 1) + rsEstatistica!QtdCapa
                Matriz(24, 2) = Matriz(24, 2) + rsEstatistica!QtdDoc
            Case "Z" 'Em   Correção de Ag. Conta
                Matriz(25, 1) = Matriz(25, 1) + rsEstatistica!QtdCapa
                Matriz(25, 2) = Matriz(25, 2) + rsEstatistica!QtdDoc
            Case "W" 'Em   Estorno
                Matriz(26, 1) = Matriz(26, 1) + rsEstatistica!QtdCapa
                Matriz(26, 2) = Matriz(26, 2) + rsEstatistica!QtdDoc
            Case "N" 'Para CSP
                Matriz(27, 1) = Matriz(27, 1) + rsEstatistica!QtdCapa
                Matriz(27, 2) = Matriz(27, 2) + rsEstatistica!QtdDoc
            Case "Q" 'Em   CSP
                Matriz(28, 1) = Matriz(28, 1) + rsEstatistica!QtdCapa
                Matriz(28, 2) = Matriz(28, 2) + rsEstatistica!QtdDoc

        End Select
        
        rsEstatistica.MoveNext
    Wend
    
    rsEstatistica.Close
End Sub
Private Sub AtualizaGrafico(ByVal IdEnv_Mal As String)
    Dim CountCapa As Long
    Dim CountDocto As Long
    Dim Count As Integer
    Dim TotPorCapa, TotPorDoc As Double
    
    
    tmrAtualiza.Enabled = False

    lblData.Caption = Format(Now, "dd/mm/yyyy")
    lblHora.Caption = Format(Now, "hh:mm:ss")
    
    CountCapa = ObtemTotalCapa(IdEnv_Mal)
    If CountCapa = -1 Then
        Exit Sub
    End If
    
    CountDocto = ObtemTotalDocumento(IdEnv_Mal)
    If CountDocto = -1 Then
        Exit Sub
    End If
    
    If Not ObtemEstatistica(IdEnv_Mal) Then
        Exit Sub
    End If
    
    Preenche_Matriz
    
    Grafico.Visible = False
    
    TotPorCapa = 0
    TotPorDoc = 0
    
    Grafico.ColumnCount = 28
    For Count = 1 To 28
        lblQtdCapa(Count - 1).Caption = FormataQuantidade(Matriz(Count, 1))
        lblQtdDoc(Count - 1).Caption = FormataQuantidade(Matriz(Count, 2))
        
        Grafico.Column = Count
        If CountCapa > 0 Then
            lblPorCapa(Count - 1).Caption = Format((Matriz(Count, 1) * 100 / CountCapa), "0.0")
            TotPorCapa = TotPorCapa + (Matriz(Count, 1) * 100 / CountCapa)
            Grafico.Data = (Matriz(Count, 1) * 100 / CountCapa)
        Else
            lblPorCapa(Count - 1).Caption = "0.0"
            Grafico.Data = 0
        End If
        
        If CountDocto > 0 Then
            lblPorDoc(Count - 1).Caption = Format((Matriz(Count, 2) * 100 / CountDocto), "0.0")
            TotPorDoc = TotPorDoc + (Matriz(Count, 2) * 100 / CountDocto)
        Else
            lblPorDoc(Count - 1).Caption = "0.0"
        End If
    Next
    
    lblTotCapa.Caption = FormataQuantidade(CountCapa)
    lblTotPorCapa.Caption = Format(TotPorCapa, "0.0")
    lblTotDoc.Caption = FormataQuantidade(CountDocto)
    lblTotPorDoc.Caption = Format(TotPorDoc, "0.0")
    
    Grafico.Visible = True
    
    tmrAtualiza.Enabled = True

End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    On Error GoTo ERRO_IMPRESSAO
    PrintForm
    On Error GoTo 0
    Exit Sub
ERRO_IMPRESSAO:
    MsgBox "Erro na impressão da Estatística.", vbCritical + vbOKOnly, App.Title
End Sub
Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(12)
   
   tmrAtualiza.Interval = Geral.Atualizacao * 1000
   tmrAtualiza.Enabled = True
   
  'Agência Processadora
   LblAgProc.Caption = Geral.AgenciaCentral
   
  'Dataprocessamento
   lblDataProc.Caption = Mid(Geral.DataProcessamento, 7, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 5, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 1, 4)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Count As Integer
    
    Grafico.ColumnCount = 26
    If KeyCode = vbKeyF2 Then
        For Count = 1 To 26
            Grafico.Column = Count
            Grafico.Data = 5.55
        Next
    End If
End Sub
Private Sub Form_Load()
    Set qryGetEstatistica = Geral.Banco.CreateQuery("", "{Call GetEstatistica (?,?)}")
    Set qryGetTotalCapa = Geral.Banco.CreateQuery("", "{Call GetTotalCapa (?,?)}")
    Set qryGetTotalDocumento = Geral.Banco.CreateQuery("", "{Call GetTotalDocumento (?,?)}")
     
End Sub
Private Sub Form_Unload(Cancel As Integer)
    tmrAtualiza.Enabled = False
    
    qryGetEstatistica.Close
    qryGetTotalCapa.Close
    qryGetTotalDocumento.Close
End Sub



Private Sub lblCor_DblClick(Index As Integer)
    
    ListaCapa.m_idtela = 0
    
    If Matriz(Index + 1, 1) = 0 Then
        Exit Sub
    End If
    
    tmrAtualiza.Enabled = False
    
    Load ListaCapa
    
    If optFiltro(0).Value Then
        ListaCapa.m_IdEnv_Mal = "T"
    ElseIf optFiltro(1).Value Then
        ListaCapa.m_IdEnv_Mal = "E"
    ElseIf optFiltro(2).Value Then
        ListaCapa.m_IdEnv_Mal = "M"
    Else
        ListaCapa.m_IdEnv_Mal = "F"
    End If
    
    Select Case Index
        Case 0
            ListaCapa.m_InStatus = "'0'"
        Case 1
            ListaCapa.m_InStatus = "'1'"
        Case 2
            ListaCapa.m_InStatus = "'2'"
        Case 3
            ListaCapa.m_InStatus = "'8','9'"
        Case 4
            ListaCapa.m_InStatus = "'7'"
        Case 5
            ListaCapa.m_InStatus = "'J'"
        Case 6
            ListaCapa.m_InStatus = "'4'"
        Case 7
            ListaCapa.m_InStatus = "'G'"
        Case 8
            ListaCapa.m_InStatus = "'5'"
        Case 9
            ListaCapa.m_InStatus = "'H'"
        Case 10
            ListaCapa.m_InStatus = "'O'"
        Case 11
            ListaCapa.m_InStatus = "'6'"
        Case 12
            ListaCapa.m_InStatus = "'I'"
        Case 13
            ListaCapa.m_InStatus = "'P','R'"
        Case 14
            ListaCapa.m_InStatus = "'S'"
        Case 15
            ListaCapa.m_InStatus = "'T'"
        Case 16
            ListaCapa.m_InStatus = "'K'"
        Case 17
            ListaCapa.m_InStatus = "'E'"
        Case 18
            ListaCapa.m_InStatus = "'D','F','X'"
        Case 19
            ListaCapa.m_InStatus = "'A'"
        Case 20
            ListaCapa.m_InStatus = "'B'"
        Case 21
            ListaCapa.m_InStatus = "'L'"
        Case 22
            ListaCapa.m_InStatus = "'M'"
        Case 23
            ListaCapa.m_InStatus = "'Y'"
        Case 24
            ListaCapa.m_InStatus = "'Z'"
        Case 25
            ListaCapa.m_InStatus = "'W'"
        Case 26
            ListaCapa.m_InStatus = "'N'"
        Case 27
            ListaCapa.m_InStatus = "'Q'"
    
    End Select
    
    ListaCapa.Caption = lblStatus(Index).Caption
    ListaCapa.Show vbModal, Me
    
    tmrAtualiza.Enabled = True
    
End Sub



Private Sub optFiltro_Click(Index As Integer)
      
    '* Colocar ampulheta antes das atualizações *'
    Screen.MousePointer = vbHourglass
    
        If optFiltro(0).Value Then
            lblFiltro.Caption = "Todos"
            AtualizaGrafico ("T")
        ElseIf optFiltro(1).Value Then
            lblFiltro.Caption = "Env."
            AtualizaGrafico ("E")
        ElseIf optFiltro(2).Value Then
            lblFiltro.Caption = "Malotes"
            AtualizaGrafico ("M")
        Else
            lblFiltro.Caption = "Fininv."
            AtualizaGrafico ("F")
        End If
            
    '* Colocar Ampulheta default depois das atualizações *'
    Screen.MousePointer = vbDefault

End Sub
Private Sub tmrAtualiza_Timer()
    If optFiltro(0).Value Then
        AtualizaGrafico ("T")
    ElseIf optFiltro(1).Value Then
        AtualizaGrafico ("E")
    Else
        AtualizaGrafico ("M")
    End If
End Sub
