VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Estatistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Captura - Estatística"
   ClientHeight    =   6096
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   10992
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6096
   ScaleWidth      =   10992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tmr_Hora 
      Interval        =   1000
      Left            =   1944
      Top             =   5640
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   384
      Left            =   5520
      TabIndex        =   7
      Top             =   5640
      Width           =   1512
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   384
      Left            =   3960
      TabIndex        =   6
      Top             =   5640
      Width           =   1512
   End
   Begin VB.PictureBox Picture5 
      Height          =   300
      Left            =   1056
      ScaleHeight     =   252
      ScaleWidth      =   2064
      TabIndex        =   4
      Top             =   144
      Width           =   2112
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Data do Movimento"
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
         Height          =   192
         Left            =   24
         TabIndex        =   5
         Top             =   24
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   300
      Left            =   3180
      ScaleHeight     =   252
      ScaleWidth      =   1212
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   144
      Width           =   1260
      Begin VB.Label lblDataProc 
         Alignment       =   2  'Center
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
         Height          =   204
         Left            =   48
         TabIndex        =   3
         Top             =   24
         Width           =   1152
      End
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   5172
      Left            =   192
      OleObjectBlob   =   "Estatistica.frx":0000
      TabIndex        =   93
      Top             =   120
      Width           =   5196
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
      Height          =   264
      Left            =   8184
      TabIndex        =   98
      Top             =   5040
      Width           =   732
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
      Height          =   264
      Left            =   8928
      TabIndex        =   97
      Top             =   5040
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
      Height          =   264
      Left            =   9492
      TabIndex        =   96
      Top             =   5040
      Width           =   876
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
      Height          =   264
      Left            =   10392
      TabIndex        =   95
      Top             =   5040
      Width           =   492
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
      Height          =   264
      Left            =   6072
      TabIndex        =   94
      Top             =   5040
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   21
      Left            =   5640
      TabIndex        =   1
      Top             =   3552
      Width           =   408
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rejeitado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   21
      Left            =   6084
      TabIndex        =   92
      Top             =   3552
      Width           =   2076
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
      Height          =   264
      Index           =   10
      Left            =   8952
      TabIndex        =   91
      Top             =   3552
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
      Height          =   264
      Index           =   10
      Left            =   9492
      TabIndex        =   90
      Top             =   3552
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
      Height          =   264
      Index           =   10
      Left            =   10392
      TabIndex        =   89
      Top             =   3552
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
      Height          =   264
      Index           =   10
      Left            =   8184
      TabIndex        =   88
      Top             =   3552
      Width           =   732
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   0
      Left            =   5640
      TabIndex        =   87
      Top             =   4224
      Width           =   408
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Excluído"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   6072
      TabIndex        =   86
      Top             =   4224
      Width           =   2076
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
      Height          =   264
      Index           =   12
      Left            =   8184
      TabIndex        =   85
      Top             =   4224
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
      Height          =   264
      Index           =   12
      Left            =   8952
      TabIndex        =   84
      Top             =   4224
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
      Height          =   264
      Index           =   12
      Left            =   9492
      TabIndex        =   83
      Top             =   4224
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
      Height          =   264
      Index           =   12
      Left            =   10392
      TabIndex        =   82
      Top             =   4224
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
      Height          =   264
      Index           =   11
      Left            =   8184
      TabIndex        =   81
      Top             =   3888
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
      Height          =   264
      Index           =   11
      Left            =   10392
      TabIndex        =   80
      Top             =   3888
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
      Height          =   264
      Index           =   11
      Left            =   9492
      TabIndex        =   79
      Top             =   3888
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
      Height          =   264
      Index           =   11
      Left            =   8952
      TabIndex        =   78
      Top             =   3888
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirmado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   17
      Left            =   6072
      TabIndex        =   77
      Top             =   3888
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   17
      Left            =   5640
      TabIndex        =   76
      Top             =   3888
      Width           =   408
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
      Height          =   264
      Index           =   9
      Left            =   8184
      TabIndex        =   75
      Top             =   3216
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
      Height          =   264
      Index           =   9
      Left            =   10392
      TabIndex        =   74
      Top             =   3216
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
      Height          =   264
      Index           =   9
      Left            =   9492
      TabIndex        =   73
      Top             =   3216
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
      Height          =   264
      Index           =   9
      Left            =   8952
      TabIndex        =   72
      Top             =   3216
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corrigido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   4
      Left            =   6072
      TabIndex        =   71
      Top             =   3216
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   4
      Left            =   5640
      TabIndex        =   70
      Top             =   3216
      Width           =   408
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
      Height          =   264
      Index           =   8
      Left            =   8184
      TabIndex        =   69
      Top             =   2880
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
      Height          =   264
      Index           =   8
      Left            =   10392
      TabIndex        =   68
      Top             =   2880
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
      Height          =   264
      Index           =   8
      Left            =   9492
      TabIndex        =   67
      Top             =   2880
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
      Height          =   264
      Index           =   8
      Left            =   8952
      TabIndex        =   66
      Top             =   2880
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transmitido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   15
      Left            =   6072
      TabIndex        =   65
      Top             =   2880
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   15
      Left            =   5640
      TabIndex        =   64
      Top             =   2880
      Width           =   408
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
      Height          =   264
      Index           =   1
      Left            =   8184
      TabIndex        =   62
      Top             =   516
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
      Height          =   264
      Index           =   1
      Left            =   10392
      TabIndex        =   61
      Top             =   516
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
      Height          =   264
      Index           =   1
      Left            =   9492
      TabIndex        =   60
      Top             =   516
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
      Height          =   264
      Index           =   1
      Left            =   8952
      TabIndex        =   59
      Top             =   516
      Width           =   492
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
      Height          =   264
      Index           =   2
      Left            =   6072
      TabIndex        =   58
      Top             =   516
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   2
      Left            =   5640
      TabIndex        =   57
      Top             =   516
      Width           =   408
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
      Height          =   264
      Index           =   2
      Left            =   8184
      TabIndex        =   56
      Top             =   852
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
      Height          =   264
      Index           =   2
      Left            =   10392
      TabIndex        =   55
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
      Height          =   264
      Index           =   2
      Left            =   9492
      TabIndex        =   54
      Top             =   852
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
      Height          =   264
      Index           =   2
      Left            =   8952
      TabIndex        =   53
      Top             =   852
      Width           =   492
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
      Height          =   264
      Index           =   6
      Left            =   6072
      TabIndex        =   52
      Top             =   852
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   6
      Left            =   5640
      TabIndex        =   51
      Top             =   852
      Width           =   408
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
      Height          =   264
      Index           =   3
      Left            =   8184
      TabIndex        =   50
      Top             =   1188
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
      Height          =   264
      Index           =   3
      Left            =   10392
      TabIndex        =   49
      Top             =   1188
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
      Height          =   264
      Index           =   3
      Left            =   9492
      TabIndex        =   48
      Top             =   1188
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
      Height          =   264
      Index           =   3
      Left            =   8952
      TabIndex        =   47
      Top             =   1188
      Width           =   492
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
      Height          =   264
      Index           =   7
      Left            =   6072
      TabIndex        =   46
      Top             =   1200
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   7
      Left            =   5640
      TabIndex        =   45
      Top             =   1188
      Width           =   408
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
      Height          =   264
      Index           =   4
      Left            =   8184
      TabIndex        =   44
      Top             =   1524
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
      Height          =   264
      Index           =   4
      Left            =   10392
      TabIndex        =   43
      Top             =   1524
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
      Height          =   264
      Index           =   4
      Left            =   9492
      TabIndex        =   42
      Top             =   1524
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
      Height          =   264
      Index           =   4
      Left            =   8952
      TabIndex        =   41
      Top             =   1524
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   8
      Left            =   6072
      TabIndex        =   40
      Top             =   1524
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   8
      Left            =   5640
      TabIndex        =   39
      Top             =   1524
      Width           =   408
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
      Height          =   264
      Index           =   5
      Left            =   8184
      TabIndex        =   38
      Top             =   1860
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
      Height          =   264
      Index           =   5
      Left            =   10392
      TabIndex        =   37
      Top             =   1860
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
      Height          =   264
      Index           =   5
      Left            =   9492
      TabIndex        =   36
      Top             =   1860
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
      Height          =   264
      Index           =   5
      Left            =   8952
      TabIndex        =   35
      Top             =   1860
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Em Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   9
      Left            =   6072
      TabIndex        =   34
      Top             =   1860
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   9
      Left            =   5640
      TabIndex        =   33
      Top             =   1860
      Width           =   408
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
      Height          =   264
      Index           =   6
      Left            =   8184
      TabIndex        =   32
      Top             =   2196
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
      Height          =   264
      Index           =   6
      Left            =   10392
      TabIndex        =   31
      Top             =   2196
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
      Height          =   264
      Index           =   6
      Left            =   9492
      TabIndex        =   30
      Top             =   2196
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
      Height          =   264
      Index           =   6
      Left            =   8952
      TabIndex        =   29
      Top             =   2196
      Width           =   492
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
      Height          =   264
      Index           =   13
      Left            =   6072
      TabIndex        =   28
      Top             =   2196
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   13
      Left            =   5640
      TabIndex        =   27
      Top             =   2196
      Width           =   408
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
      Height          =   264
      Index           =   7
      Left            =   8184
      TabIndex        =   26
      Top             =   2544
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
      Height          =   264
      Index           =   7
      Left            =   10392
      TabIndex        =   25
      Top             =   2544
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
      Height          =   264
      Index           =   7
      Left            =   9492
      TabIndex        =   24
      Top             =   2544
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
      Height          =   264
      Index           =   7
      Left            =   8952
      TabIndex        =   23
      Top             =   2544
      Width           =   492
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
      Height          =   264
      Index           =   14
      Left            =   6072
      TabIndex        =   22
      Top             =   2544
      Width           =   2076
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   14
      Left            =   5640
      TabIndex        =   21
      Top             =   2544
      Width           =   408
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
      Height          =   264
      Left            =   6072
      TabIndex        =   20
      Top             =   168
      Width           =   2076
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
      Height          =   264
      Left            =   8184
      TabIndex        =   19
      Top             =   168
      Width           =   732
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
      Height          =   264
      Left            =   8952
      TabIndex        =   18
      Top             =   168
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
      Height          =   264
      Left            =   9492
      TabIndex        =   17
      Top             =   168
      Width           =   876
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
      Height          =   264
      Left            =   10392
      TabIndex        =   16
      Top             =   168
      Width           =   492
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   18
      Left            =   5544
      TabIndex        =   15
      Top             =   7500
      Width           =   408
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
      Height          =   264
      Index           =   18
      Left            =   5976
      TabIndex        =   14
      Top             =   7500
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
      Height          =   264
      Index           =   18
      Left            =   8616
      TabIndex        =   13
      Top             =   7500
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
      Height          =   264
      Index           =   18
      Left            =   9384
      TabIndex        =   12
      Top             =   7500
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
      Height          =   264
      Index           =   18
      Left            =   9924
      TabIndex        =   11
      Top             =   7500
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
      Height          =   264
      Index           =   18
      Left            =   10824
      TabIndex        =   10
      Top             =   7500
      Width           =   492
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17/07/2000"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   48
      TabIndex        =   9
      Top             =   5700
      Width           =   936
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:30:35"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1044
      TabIndex        =   8
      Top             =   5700
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   5352
      Left            =   48
      TabIndex        =   0
      Top             =   72
      Width           =   5448
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   5352
      Left            =   5568
      TabIndex        =   63
      Top             =   72
      Width           =   5400
   End
End
Attribute VB_Name = "Estatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Matriz(1 To 12, 1 To 2) As Long
Private Sub CmdFechar_Click()

    Unload Me
End Sub


Private Sub Preenche_Matriz()

    Dim Selecionar          As New Custodia.Selecionar
    Dim RsTotalBordero      As New ADODB.Recordset
    Dim RsTotalCheques      As New ADODB.Recordset
    Dim RsTotalCapa         As New ADODB.Recordset
    Dim RsTotalDocto        As New ADODB.Recordset

    Dim Count               As Integer
    Dim CountCapa           As Integer
    Dim TotPorCapa          As Integer
    Dim CountDocto          As Integer
    Dim TotPorDoc           As Integer

    On Error GoTo Preenche_Matriz_Err

    'Calcular a quantidade total de capas da data
    Set RsTotalCapa = g_cMainConnection.Execute(Selecionar.GetTotalCapa(Geral.DataProcessamento))
    If Not RsTotalCapa.EOF Then
        CountCapa = RsTotalCapa!Totalcapa
    Else
        MsgBox "Não foi possível exibir gráfico", vbInformation + vbOKOnly
        Call Zera_Matriz
        Exit Sub
    End If

    If CountCapa = 0 Then
        'Desabilita gráfico
        Grafico.Visible = False
        Exit Sub
    End If

    'Calcular a quantidade total de cheques da data
    Set RsTotalDocto = g_cMainConnection.Execute(Selecionar.GetTotalDocto(Geral.DataProcessamento))
    If Not RsTotalDocto.EOF Then
        CountDocto = RsTotalDocto!TotalDocto
    Else
        MsgBox "Não foi possível exibir gráfico", vbInformation + vbOKOnly
        Call Zera_Matriz
        Exit Sub
    End If

    'Selecionar totais do bordero para preencher gráfico
    Set RsTotalBordero = g_cMainConnection.Execute(Selecionar.GetEstatisticaBordero(Geral.DataProcessamento))

    While Not RsTotalBordero.EOF
        Select Case RsTotalBordero!Status
            Case "2" 'Em Complementacao
                Matriz(1, 1) = Matriz(1, 1) + RsTotalBordero!QtdBordero
            Case "4" 'Para Prova Zero
                Matriz(2, 1) = Matriz(2, 1) + RsTotalBordero!QtdBordero
            Case "G" 'Em Prova Zero
                Matriz(3, 1) = Matriz(3, 1) + RsTotalBordero!QtdBordero
            Case "5" 'Para Supervisor
                Matriz(4, 1) = Matriz(4, 1) + RsTotalBordero!QtdBordero
            Case "H" 'Em Supervisor
                Matriz(5, 1) = Matriz(5, 1) + RsTotalBordero!QtdBordero
            Case "R" 'Para Transmissao
                Matriz(6, 1) = Matriz(6, 1) + RsTotalBordero!QtdBordero
            Case "S" 'Em Transmissao
                Matriz(7, 1) = Matriz(7, 1) + RsTotalBordero!QtdBordero
            Case "T" 'Transmitido
                Matriz(8, 1) = Matriz(8, 1) + RsTotalBordero!QtdBordero
            Case "C" 'Corrigido
                Matriz(9, 1) = Matriz(9, 1) + RsTotalBordero!QtdBordero
            Case "X" 'Rejeitado
                Matriz(10, 1) = Matriz(10, 1) + RsTotalBordero!QtdBordero
            Case "E" 'Confimado
                Matriz(11, 1) = Matriz(11, 1) + RsTotalBordero!QtdBordero
            Case "D" 'Excluido
                Matriz(12, 1) = Matriz(12, 1) + RsTotalBordero!QtdBordero
        End Select

        RsTotalBordero.MoveNext
    Wend

    'Selecionar totais do cheque para preencher gráfico
    Set RsTotalCheques = g_cMainConnection.Execute(Selecionar.GetEstatisticaCheque(Geral.DataProcessamento))

    While Not RsTotalCheques.EOF
        Select Case RsTotalCheques!Status
            Case "2" 'Em Complementacao
                Matriz(1, 2) = Matriz(1, 2) + RsTotalCheques!QtdCheque
            Case "4" 'Para Prova Zero
                Matriz(2, 2) = Matriz(2, 2) + RsTotalCheques!QtdCheque
            Case "G" 'Em Prova Zero
                Matriz(3, 2) = Matriz(3, 2) + RsTotalCheques!QtdCheque
            Case "5" 'Para Supervisor
                Matriz(4, 2) = Matriz(4, 2) + RsTotalCheques!QtdCheque
            Case "H" 'Em Supervisor
                Matriz(5, 2) = Matriz(5, 2) + RsTotalCheques!QtdCheque
            Case "R" 'Para Transmissao
                Matriz(6, 2) = Matriz(6, 2) + RsTotalCheques!QtdCheque
            Case "S" 'Em Transmissao
                Matriz(7, 2) = Matriz(7, 2) + RsTotalCheques!QtdCheque
            Case "T" 'Transmitido
                Matriz(8, 2) = Matriz(8, 2) + RsTotalCheques!QtdCheque
            Case "C" 'Corrigido
                Matriz(9, 2) = Matriz(9, 2) + RsTotalCheques!QtdCheque
            Case "X" 'Rejeitado
                Matriz(10, 2) = Matriz(10, 2) + RsTotalCheques!QtdCheque
            Case "E" 'Confirmado
                Matriz(11, 2) = Matriz(11, 2) + RsTotalCheques!QtdCheque
            Case "D" 'Excluido
                Matriz(12, 2) = Matriz(12, 2) + RsTotalCheques!QtdCheque
        End Select

        RsTotalCheques.MoveNext
    Wend

    Grafico.ColumnCount = 12
    For Count = 1 To 12
        lblQtdCapa(Count).Caption = FormataQuantidade(Matriz(Count, 1))
        lblQtdDoc(Count).Caption = FormataQuantidade(Matriz(Count, 2))

        Grafico.Column = Count
        If CountCapa > 0 Then
            lblPorCapa(Count).Caption = Format((Matriz(Count, 1) * 100 / CountCapa), "0.0")
            TotPorCapa = TotPorCapa + (Matriz(Count, 1) * 100 / CountCapa)
            Grafico.Data = (Matriz(Count, 1) * 100 / CountCapa)
        Else
            lblPorCapa(Count).Caption = "0.0"
            Grafico.Data = 0
        End If

        If CountDocto > 0 Then
            lblPorDoc(Count).Caption = Format((Matriz(Count, 2) * 100 / CountDocto), "0.0")
            TotPorDoc = TotPorDoc + (Matriz(Count, 2) * 100 / CountDocto)
        Else
            lblPorDoc(Count).Caption = "0.0"
        End If
    Next

    lblTotCapa.Caption = FormataQuantidade(CountCapa)
    lblTotPorCapa.Caption = Format(TotPorCapa, "0.0")
    lblTotDoc.Caption = FormataQuantidade(CountDocto)
    lblTotPorDoc.Caption = Format(TotPorDoc, "0.0")

    'Fechando os recordsets
    RsTotalBordero.Close
    RsTotalCheques.Close
    RsTotalCapa.Close

    Exit Sub

Preenche_Matriz_Err:
    Call TratamentoErro("Erro ao pesquisar Capas para montar o Gráfico.", Err)
    Unload Me
End Sub

Private Function FormataQuantidade(ByVal Qtd As Long) As String
    Dim strValor As String
    Dim strResult As String
    Dim Count As Integer
    
    strValor = Trim(Str(Qtd))
    
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

Private Sub Zera_Matriz()

    Dim Count As Integer

    For Count = 1 To 12
        Matriz(Count, 1) = 0
        Matriz(Count, 2) = 0
        lblQtdCapa(Count).Caption = "0"
        lblQtdDoc(Count).Caption = "0"
        lblPorCapa(Count).Caption = "0"
        lblPorDoc(Count).Caption = "0"
    Next

    lblTotCapa.Caption = "0"
    lblTotPorCapa.Caption = "0"
    lblTotDoc.Caption = "0"
    lblTotPorDoc.Caption = "0"
    
End Sub

Private Sub cmdImprimir_Click()

    Me.PrintForm
End Sub
Private Sub Form_Activate()

    Call Zera_Matriz
    
    Call Preenche_Matriz
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim Count As Integer
    
    Grafico.ColumnCount = 12
    If KeyCode = vbKeyF2 Then
        For Count = 1 To 12
            Grafico.Column = Count
            Grafico.Data = 5.55
        Next
    End If
End Sub

Private Sub Form_Load()

    lblData.Caption = Format(Now, "dd/mm/yyyy")
    lblHora.Caption = Format(Now, "hh:mm:ss")
    lblDataProc.Caption = Format(Format(Geral.DataProcessamento, "0000/00/00"), "dd/mm/yyyy")

End Sub

Private Sub Tmr_Hora_Timer()

    lblData.Caption = Format(Now, "dd/mm/yyyy")
    lblHora.Caption = Format(Now, "hh:mm:ss")
End Sub
