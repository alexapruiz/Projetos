VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCSP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.S.P."
   ClientHeight    =   8604
   ClientLeft      =   156
   ClientTop       =   372
   ClientWidth     =   12072
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8604
   ScaleWidth      =   12072
   Begin VB.Frame fraSaldo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3252
      Left            =   1800
      TabIndex        =   51
      Top             =   5040
      Visible         =   0   'False
      Width           =   6372
      Begin VB.CommandButton cmd_Sld_Fechar 
         Caption         =   "&Fechar"
         Height          =   372
         Left            =   2280
         TabIndex        =   61
         Top             =   2760
         Width           =   1812
      End
      Begin VB.Frame fra_Sld_Saldo 
         Height          =   1692
         Left            =   360
         TabIndex        =   55
         Top             =   960
         Width           =   5532
         Begin VB.Label lbl_Sld_SaldoDisponivel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   65
            Top             =   1320
            Width           =   2172
         End
         Begin VB.Label lbl_Sld_ValorBloqueado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   64
            Top             =   960
            Width           =   2172
         End
         Begin VB.Label lbl_Sld_LimiteCheque 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   63
            Top             =   600
            Width           =   2172
         End
         Begin VB.Label lbl_Sld_DataSaldo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   62
            Top             =   240
            Width           =   2172
         End
         Begin VB.Label lbl_Sld_SaldoDisponivel 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponível:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   1548
         End
         Begin VB.Label lbl_Sld_ValorBloqueado 
            AutoSize        =   -1  'True
            Caption         =   "Valor Bloqueado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Top             =   960
            Width           =   1524
         End
         Begin VB.Label lbl_Sld_LimiteCheque 
            AutoSize        =   -1  'True
            Caption         =   "Limite de Cheque Especial:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   600
            Width           =   2424
         End
         Begin VB.Label lbl_Sld_DataSaldo 
            AutoSize        =   -1  'True
            Caption         =   "Data/Hora do Saldo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   1812
         End
      End
      Begin VB.Frame fra_Sld_Cheque 
         Height          =   612
         Left            =   360
         TabIndex        =   53
         Top             =   240
         Width           =   5532
         Begin VB.Label lbl_Sld_ValorCheque 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   60
            Top             =   192
            Width           =   2172
         End
         Begin VB.Label lbl_Sld_ValorCheque 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Cheque:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   54
            Top             =   240
            Width           =   1524
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Consulta de saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   216
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   6612
      End
   End
   Begin VB.PictureBox frmLocalizar 
      Height          =   1272
      Left            =   4440
      ScaleHeight     =   1224
      ScaleWidth      =   2604
      TabIndex        =   44
      Top             =   1962
      Visible         =   0   'False
      Width           =   2652
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1464
         TabIndex        =   48
         Top             =   816
         Width           =   972
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "&Localizar"
         Height          =   300
         Left            =   144
         TabIndex        =   46
         Top             =   816
         Width           =   972
      End
      Begin VB.TextBox txtNumEnvMal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   18
         TabIndex        =   45
         Top             =   384
         Width           =   2304
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número Envelope/Malote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   47
         Top             =   96
         Width           =   2232
      End
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2820
      ScaleHeight     =   1884
      ScaleWidth      =   5724
      TabIndex        =   32
      Top             =   2976
      Visible         =   0   'False
      Width           =   5772
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2184
         TabIndex        =   40
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   336
         TabIndex        =   41
         Top             =   912
         Width           =   4932
         _ExtentX        =   8700
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos enviados para C.S.P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   348
         TabIndex        =   42
         Top             =   576
         Width           =   4560
      End
   End
   Begin VB.Timer TmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   9756
      Top             =   24
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   24
      ScaleHeight     =   216
      ScaleWidth      =   1752
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   48
      Width           =   1800
      Begin VB.Label LblEnv_Mal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   228
         Left            =   108
         TabIndex        =   33
         Top             =   0
         Width           =   1296
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   252
      Left            =   1896
      ScaleHeight     =   204
      ScaleWidth      =   8160
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   384
      Width           =   8208
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Estorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   3360
         TabIndex        =   66
         Top             =   0
         Width           =   648
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   2640
         TabIndex        =   50
         Top             =   0
         Width           =   636
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Duplic."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   1968
         TabIndex        =   39
         Top             =   0
         Width           =   588
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   84
         TabIndex        =   29
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   600
         TabIndex        =   28
         Top             =   0
         Width           =   624
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ocorr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   1368
         TabIndex        =   27
         Top             =   0
         Width           =   516
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   4140
         TabIndex        =   26
         Top             =   0
         Width           =   948
      End
      Begin VB.Label Label6 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   216
         Left            =   7416
         TabIndex        =   25
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox PctMalote 
      Height          =   264
      Left            =   3828
      ScaleHeight     =   216
      ScaleWidth      =   1176
      TabIndex        =   22
      Top             =   48
      Width           =   1224
      Begin VB.Label Label11 
         Caption         =   "Nro. Malote"
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
         Height          =   228
         Left            =   36
         TabIndex        =   23
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   264
      Left            =   1884
      ScaleHeight     =   216
      ScaleWidth      =   528
      TabIndex        =   20
      Top             =   48
      Width           =   576
      Begin VB.Label Label12 
         Caption         =   "Lote"
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
         Height          =   228
         Left            =   24
         TabIndex        =   21
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame FraCmd 
      Height          =   4260
      Left            =   10164
      TabIndex        =   19
      Top             =   -72
      Width           =   1872
      Begin VB.CommandButton cmdEnviarExpedicao 
         Caption         =   "Enviar &Expedição"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   1584
      End
      Begin VB.CommandButton cmdRemoverVinculo 
         Caption         =   "Re&mover Vínculo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2016
         Width           =   1584
      End
      Begin VB.CommandButton cmdEstorno 
         Caption         =   "&Estorno"
         Height          =   300
         Left            =   -336
         TabIndex        =   10
         Top             =   4128
         Visible         =   0   'False
         Width           =   1584
      End
      Begin VB.CommandButton cmdSaldo 
         Caption         =   "&Saldo"
         Height          =   300
         Left            =   144
         TabIndex        =   7
         Top             =   2760
         Width           =   1584
      End
      Begin VB.CommandButton cmdVincular 
         Caption         =   "&Vincular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1656
         Width           =   1584
      End
      Begin VB.CommandButton cmdRetiraOcorrencia 
         Caption         =   "&Retirar Ocorrência"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1296
         Width           =   1584
      End
      Begin VB.CommandButton CmdEnviaTransmissao 
         Caption         =   "Enviar &Transmissâo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3504
         Width           =   1584
      End
      Begin VB.CommandButton CmdLocalizar 
         Caption         =   "&Localizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   576
         Width           =   1584
      End
      Begin VB.CommandButton CmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   300
         Left            =   144
         TabIndex        =   0
         Top             =   216
         Width           =   1584
      End
      Begin VB.CommandButton CmdEnviaCompensacao 
         Caption         =   "Enviar &Compensação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   1584
      End
      Begin VB.CommandButton CmdOcorrencia 
         Caption         =   "&Ocorrência"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   936
         Width           =   1584
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   144
         TabIndex        =   11
         Top             =   3864
         Width           =   1584
      End
   End
   Begin VB.Frame FrmCmdImagem 
      Height          =   4368
      Left            =   10164
      TabIndex        =   18
      Top             =   4164
      Width           =   1884
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   528
         Picture         =   "frmCSP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   372
         Width           =   888
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   528
         Picture         =   "frmCSP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1140
         Width           =   888
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   528
         Picture         =   "frmCSP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1896
         Width           =   888
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverte cor"
         Height          =   696
         Left            =   528
         Picture         =   "frmCSP.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2664
         Width           =   888
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   528
         Picture         =   "frmCSP.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3420
         Width           =   888
      End
   End
   Begin VB.ListBox lstCapa 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2904
      ItemData        =   "frmCSP.frx":0F32
      Left            =   24
      List            =   "frmCSP.frx":0F34
      TabIndex        =   34
      Top             =   456
      Width           =   1800
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4740
      Left            =   24
      TabIndex        =   17
      Top             =   3792
      Width           =   10080
      Begin LeadLib.Lead Lead1 
         Height          =   4452
         Left            =   96
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   216
         Width           =   9828
         _Version        =   524288
         _ExtentX        =   17336
         _ExtentY        =   7853
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   369
         ScaleWidth      =   817
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Timer TmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9360
      Top             =   24
   End
   Begin VB.ListBox LstDocto 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "frmCSP.frx":0F36
      Left            =   1884
      List            =   "frmCSP.frx":0F38
      MultiSelect     =   2  'Extended
      TabIndex        =   31
      Top             =   672
      Width           =   8220
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   2700
      Left            =   1872
      ScaleHeight     =   2652
      ScaleWidth      =   8172
      TabIndex        =   49
      Top             =   672
      Width           =   8220
   End
   Begin VB.Label lblAjuste 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1000
      Left            =   0
      TabIndex        =   67
      Top             =   3744
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label LblCapaDup 
      Caption         =   "Capa Duplicada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   6744
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblOcorrencia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ocorrência:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   276
      Left            =   24
      TabIndex        =   37
      Top             =   3504
      Width           =   10080
   End
   Begin VB.Label lblNumMalote 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   5136
      TabIndex        =   36
      Top             =   48
      Width           =   1500
   End
   Begin VB.Label lblLote 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2520
      TabIndex        =   35
      Top             =   48
      Width           =   1224
   End
End
Attribute VB_Name = "frmCSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Delaração dos Objetos RDO
Private qryGetCapa                      As rdoQuery
Private qryGetDocumentos                As rdoQuery
Private qryAtualizaStatusCapa           As rdoQuery
Private qryGetOcorr                     As rdoQuery
Private qryAtualizaOcorrencia           As rdoQuery
Private qryRemoveAjusteCapa             As rdoQuery
Private qryEnviarCompensacao            As rdoQuery
Private qryGeraVinculo                  As rdoQuery
Private qryControleCapa                 As rdoQuery
Private qryInsereAjuste                 As rdoQuery
Private qryAlteraTipoDocto              As rdoQuery
Private qryGetUltimaOrdemCaptura        As rdoQuery

'Declaração de Variáveis
Private AlterouDocto                    As Boolean
Private PrimeiraVez                     As Boolean
Private bCapaDuplicada                  As Boolean
Private teclou                          As Boolean
Private IdSelecionado                   As Long
Private sTempo                          As Integer

Dim tpAjuste  As tpMyAjuste
Private Type tpMyAjuste
    TipoDocto                           As Integer
    Vinculo                             As Long
    Agencia                             As Integer
    Conta                               As Long
    Valor                               As Currency
End Type

'Type de Capas
Private aCapa()     As TCapa
Private Type TCapa
  IdCapa                As Long
  IdLote                As Long
  IdEnv_Mal             As String * 1
  Capa                  As String * 18
  NumMalote             As String * 11
  AgOrig                As Integer
  Status                As String * 1
  Duplicidade           As Integer
End Type

'Type para Documentos
Dim bExisteAjusteEmCapa As Boolean      'Informa se existe Ajuste NÃO transmitido na Capa

Private aDoc()          As TDoc
Private Type TDoc
  NrSeq                 As Integer
  IdDocto               As Long
  IdCapa                As Long
  TipoDocto             As Integer
  DscTipoDocto          As String * 18
  Duplicidade           As Boolean
  Ocorrencia            As String * 5
  RetornoTransacao      As Long
  Leitura               As String * 48
  Frente                As String * 20
  Verso                 As String * 20
  Status                As String * 1
  Vinculo               As Long
  Valor                 As Double
  Ordem                 As String
  RetornoDeSaldo        As Boolean
  DataHoraSaldo         As String
  LimiteChequeEspecial  As String
  ValorBloqueado        As String
  SaldoDisponivel       As String
  EstornoDocto          As Boolean
  DepositoAgencia       As Integer              'Agência de Depósito ou Crédito(OCT)
  DepositoConta         As Long                 'Conta de Depósito ou Crédito(OCT)
  
End Type
Dim sMasc               As String

Sub AtualizaStatusCapa(ByVal sIdCapa As Long, sStatus As String)

  On Error GoTo ERRO_ATUALIZASTATUS

  'Se IdCapa passado = 0 , sair da funcao
  If Val(sIdCapa) = 0 Then Exit Sub

  Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusCapa (?,?,?)}")
  With qryAtualizaStatusCapa
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento 'Data Proc.
      .rdoParameters(2) = sIdCapa                 'IdCapa
      .rdoParameters(3) = sStatus                 'Status
      .Execute
  End With

  If qryAtualizaStatusCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar o status da capa.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  Exit Sub

ERRO_ATUALIZASTATUS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar o Status do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CapaDuplicada() As Boolean

  On Error GoTo ERRO_CAPADUPLICADA

  Dim sSql As String
  Dim rsCapa As rdoResultset

  CapaDuplicada = False

  If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" And aCapa(lstCapa.ListIndex + 1).AgOrig = 0 Then Exit Function

  If aCapa(lstCapa.ListIndex + 1).Duplicidade = 1 Then
    CapaDuplicada = True
  Else
    CapaDuplicada = False
  End If

  Exit Function

ERRO_CAPADUPLICADA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Duplicidade de Capas.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Private Sub FinalizaCapa(ByVal sCod As String)

    Screen.MousePointer = vbHourglass

    'Atualizar o STATUS da capa para 'R' -> Para Transmissão
    Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, sCod)
    aCapa(lstCapa.ListIndex + 1).Status = sCod

    'Limpando a variável que armazena a capa Atual
    IdSelecionado = 0
    
    Screen.MousePointer = vbDefault

    'Posicionar na próxima Capa da Lista
    Call LimpaListaDocto

    If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
      'Existem mais Capas -> Posicionar
      lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
        AlterouDocto = False
        Call CmdAtualizar_Click
      
    End If
    
End Sub
Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  TmrPesquisa.Enabled = True
  'Desabilita botões de comando
  CmdAtualizar.Enabled = False
  CmdEnviaCompensacao.Enabled = False
  CmdEnviaTransmissao.Enabled = False
  cmdEnviarExpedicao.Enabled = False
  cmdVincular.Enabled = False
'''  cmdEstorno.Enabled = False
  CmdLocalizar.Enabled = False
  CmdOcorrencia.Enabled = False
  cmdSaldo.Enabled = False
  cmdRetiraOcorrencia.Enabled = False
  cmdRemoverVinculo.Enabled = False
  Progress.Value = 0
  
  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  'Call GravaLog(0, 0, 267)
  
End Sub
Sub HDMalote(ByVal bValor As Boolean)

  PctMalote.Visible = bValor
  lblNumMalote.Visible = bValor
  If bValor = False Then
    lblLote.Caption = ""
  End If
End Sub
Sub HDObjetos(bValor As Boolean)

  On Error GoTo ERRO_HDOBJETOS

  'Habilita / Desabilita Objetos e frames para que o TAB 'TipoDocto' fique Modal
  FraCmd.Enabled = bValor
  FrmCmdImagem.Enabled = bValor
  FrmImagem.Enabled = bValor

'  lstCapa.Enabled = bValor
  LstDocto.Enabled = bValor

  Exit Sub

ERRO_HDOBJETOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Desabilitar Objetos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub HDObjetosImagem(bValor As Boolean)

  On Error GoTo ERRO_HDOBJETOS

  cmdZoomMais.Enabled = bValor
  cmdZoomMenos.Enabled = bValor
  cmdRotacao.Enabled = bValor
  cmdInverteCor.Enabled = bValor
  cmdFrenteVerso.Enabled = bValor
  FrmImagem.Visible = bValor
  Lead1.ForceRepaint

  Exit Sub

ERRO_HDOBJETOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao preparar botões de manipulação de Imagens.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Sub LimpaListaDocto()

  LstDocto.Clear
  lblOcorrencia.Caption = ""

End Sub

Sub LimpaListas()

  lstCapa.Clear
  LstDocto.Clear
  FrmImagem.Visible = False
  Erase aCapa

End Sub
Function PreencheListCapas() As Boolean

  On Error GoTo ERRO_PREENCHELISTCAPAS

  Dim rsCapa As rdoResultset
  Dim sSql As String
  Dim sPosicaoErro As String
  Dim X As Integer

  Call LimpaListas

  'Passando parâmetros para a Stored Procedure 'GetCapaParaCSP
  sSql = Geral.DataProcessamento & " , " & Geral.Intervalo

  Set qryGetCapa = Geral.Banco.CreateQuery("", "{call GetCapaParaCSP (" & sSql & ")}")

  Set rsCapa = qryGetCapa.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  If rsCapa.RowCount > 0 Then

    'Desabilitar o Timer de Pesquisa
    TmrPesquisa.Enabled = False
    FrmPesquisa.Visible = False
    'Desabilita botões de comando
    CmdAtualizar.Enabled = True
    CmdEnviaCompensacao.Enabled = True
    CmdEnviaTransmissao.Enabled = True
    cmdEnviarExpedicao.Enabled = True
    cmdVincular.Enabled = True
'''    cmdEstorno.Enabled = True
    CmdLocalizar.Enabled = True
    CmdOcorrencia.Enabled = True
    cmdSaldo.Enabled = True
    cmdRetiraOcorrencia.Enabled = True
    cmdRemoverVinculo.Enabled = True
    
    ReDim Preserve aCapa(rsCapa.RowCount)

    X = 1
    While Not rsCapa.EOF
        'Carregando o Array com as Capas
        aCapa(X).IdCapa = rsCapa!IdCapa
        aCapa(X).IdLote = rsCapa!IdLote
        aCapa(X).IdEnv_Mal = rsCapa!IdEnv_Mal
        aCapa(X).Capa = rsCapa!Capa
        aCapa(X).NumMalote = rsCapa!Num_Malote
        aCapa(X).AgOrig = rsCapa!AgOrig
        aCapa(X).Status = rsCapa!Status
        aCapa(X).Duplicidade = rsCapa!Duplicidade

        lstCapa.AddItem (rsCapa!Capa)
        lstCapa.ItemData(lstCapa.NewIndex) = rsCapa!IdCapa

        rsCapa.MoveNext
        X = X + 1
        DoEvents
    Wend
  End If

  'Selecionar a Primeira Capa , caso exista
  If lstCapa.ListCount > 0 Then
    lstCapa.Selected(0) = True
    PreencheListCapas = True
  Else
    PreencheListCapas = False
    Call HDObjetosImagem(False)
    Call HDMalote(False)
  End If

  Exit Function

ERRO_PREENCHELISTCAPAS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Ler Capas.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Function
Function CapaSelecionadaDisponivel() As Integer

On Error GoTo ERRO_CAPASELECIONADADISP

    CapaSelecionadaDisponivel = 2   'Erro
    
'   Set qryGetCapa = Geral.Banco.CreateQuery("", "{? = call VerificaCapaDisponivel (?,?,?,?,?)}")
    Set qryGetCapa = Geral.Banco.CreateQuery("", "{? = call GetCapaDisponivel (?,?,?,?,?)}")

    With qryGetCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento             'Data de Processamento
        .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa 'IdCapa
        .rdoParameters(3) = "N"                                 'Status 1
        .rdoParameters(4) = "Q"                                 'Status 2 (Pendentes)
        .rdoParameters(5) = Geral.Intervalo                     'Intervalo de Atualização
    
        .Execute
    End With

    CapaSelecionadaDisponivel = qryGetCapa(0)

    If qryGetCapa(0) <> 0 Then
        lstCapa.ListIndex = -1
        LstDocto.Clear
        FrmImagem.Visible = False
        Screen.MousePointer = vbDefault
        MsgBox "Este Envelope / Malote não está disponível. Pode estar sendo tratado por outra estação ou já foi tratado.", vbInformation, App.Title
    End If

Exit Function

ERRO_CAPASELECIONADADISP:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar se a Capa está Disponível.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PreencheListDocto(ByVal Indice As Integer)

Dim sSql As String
Dim sLinha As String
Dim rsDocumentos As rdoResultset
Dim X As Integer

On Error GoTo ERRO_PREENCHELISTDOCTO
  
    LstDocto.Visible = False
    
    'Selecionar todos os documentos pertencentes à capa selecionada
    sSql = Geral.DataProcessamento & " , " & Val(lstCapa.ItemData(lstCapa.ListIndex))
    
    Set qryGetDocumentos = Geral.Banco.CreateQuery("", "{call GetDocumentoCSP (" & sSql & ")}")
    
    Set rsDocumentos = qryGetDocumentos.OpenResultset(rdOpenStatic, rdConcurReadOnly)

    X = 1
    Call LimpaListaDocto
    ReDim aDoc(rsDocumentos.RowCount)
    bExisteAjusteEmCapa = False

    If Not rsDocumentos.EOF Then
        While Not rsDocumentos.EOF
            'Para ajuste transmitido não descartar
            If (rsDocumentos!TipoDocto = 32 Or rsDocumentos!TipoDocto = 33 Or _
                rsDocumentos!TipoDocto = 34 Or rsDocumentos!TipoDocto = 38 Or _
                rsDocumentos!TipoDocto = 42 Or rsDocumentos!TipoDocto = 43 Or _
                rsDocumentos!TipoDocto = 44 Or rsDocumentos!TipoDocto = 45) And _
                rsDocumentos!Status <> "T" Then
                rsDocumentos.MoveNext
            Else
        
                'Numero Sequencial
                aDoc(X).NrSeq = X
                sLinha = Format(aDoc(X).NrSeq, "0000") & Space(2)
        
                'Vinculo
                aDoc(X).Vinculo = Val(rsDocumentos!Vinculo & "")
                sLinha = sLinha & Format(aDoc(X).Vinculo, String(5, "0")) & Space(8 - Len(Format(aDoc(X).Vinculo, String(5, "0"))))
        
                'Ocorrencia
                aDoc(X).Ocorrencia = Val(rsDocumentos!Ocorrencia & "")
                If Val(rsDocumentos!Ocorrencia & "") <> 0 Then
                  sLinha = sLinha & "S" & Space(5)
                Else
                  sLinha = sLinha & " " & Space(5)
                End If
                
                'Retorno Transacao
                aDoc(X).RetornoTransacao = Val(rsDocumentos!RetornoTransacao)
    
                'Indicador de Duplicidade
                aDoc(X).Duplicidade = Val(rsDocumentos!Duplicidade & "")
                If Val(rsDocumentos!Duplicidade) = 1 Then
                    sLinha = sLinha & "S" & Space(6)
                Else
                    sLinha = sLinha & " " & Space(6)
                End If
        
                'Status do Documento
                aDoc(X).Status = rsDocumentos!Status & ""
                'Verifica se status docto difere dos "comuns" tratados pela CSP em
                'que não tem efeito algum, com isso será tratado como docto complementado
                If InStr("0-1-2-T-D-F-C", aDoc(X).Status) = 0 Then
                    aDoc(X).Status = "1"
                    'Muda status de documento com situação irregular (status <> 0,1,2,T,D,F e C) para "1"
                    If Not AtualizaStatusDocumento(CLng(rsDocumentos!IdDocto & ""), "1") Then
                        MsgBox "Não foi possível atualizar situação do documento (status para 1)"
                        Call LimpaListaDocto
                        ReDim aDoc(rsDocumentos.RowCount)
                        'Retorna status da capa para "N"
                        Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "N")
                        lstCapa.ListIndex = -1
                        Exit Sub
                    End If
                End If
                
                'Tipo de Documento
                aDoc(X).TipoDocto = rsDocumentos!TipoDocto & ""
    
                'Status Documento Transmitido
                If aDoc(X).Status = "T" Then
                    sLinha = sLinha & "S" & Space(5)
                Else
                    sLinha = sLinha & " " & Space(5)
                End If
    
                'Status Documento para Estorno
                aDoc(X).EstornoDocto = Not IsNull(rsDocumentos!IdDoctoEstorno)
                If aDoc(X).EstornoDocto Then
                    sLinha = sLinha & "S" & Space(4)
                Else
                    sLinha = sLinha & " " & Space(4)
                End If
                
                Select Case aDoc(X).TipoDocto
                        Case 0          'Indefinido
                            aDoc(X).DscTipoDocto = "INDEFINIDO   "
        
                        Case 1          'CAPA DE ENVELOPE / MALOTE
                            If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
                                If Val(aDoc(X).Status) = 0 Then
                                    aDoc(X).DscTipoDocto = "ENVELOPE ILEGÍVEL "
                                Else
                                    aDoc(X).DscTipoDocto = "ENVELOPE          "
                                End If
                            Else
                                If Val(aDoc(X).Status) = 0 Then
                                    aDoc(X).DscTipoDocto = "MALOTE ILEGÍVEL   "
                                Else
                                    aDoc(X).DscTipoDocto = "MALOTE            "
                                End If
                            End If
        
                        Case 2, 3       'Depósito
                            aDoc(X).DscTipoDocto = "DEPOSITO          "
        
                        Case 4          'Aut. Déb.
                          aDoc(X).DscTipoDocto = "DEBITO CC         "
                        Case 5, 6, 7    'Cheque
                            If aDoc(X).TipoDocto = 5 Then
                                aDoc(X).DscTipoDocto = "CHEQUE UBB        "
                            ElseIf aDoc(X).TipoDocto = 6 Then
                                aDoc(X).DscTipoDocto = "CHEQUE COMPENSAÇÃO"
                            ElseIf aDoc(X).TipoDocto = 7 Then
                                aDoc(X).DscTipoDocto = "CHEQUE DEPÓSITO   "
                          End If
                          
                        Case 32, 34, 42, 44     'Ajuste de Crédito
                          aDoc(X).DscTipoDocto = "AJ. CREDITO       "
                          bExisteAjusteEmCapa = True
                        Case 33, 38, 43, 45    'Ajuste de Débito
                          aDoc(X).DscTipoDocto = "AJ. DÉBITO        "
                          bExisteAjusteEmCapa = True
                        Case 36         'Cartão Avulso
                          aDoc(X).DscTipoDocto = "CARTÃO AVULSO     "
                        Case 37         'OCT
                          aDoc(X).DscTipoDocto = "OCT               "
                        Case 39         'Capa OCT
                          aDoc(X).DscTipoDocto = "CAPA OCT          "
                        Case 41         'LANÇAMENTO INTERNO
                          aDoc(X).DscTipoDocto = "LANCTO INTERNO    "
                        Case Else       'Pagamento
                          aDoc(X).DscTipoDocto = "PAGAMENTO         "
                End Select
    
                sLinha = sLinha & aDoc(X).DscTipoDocto & Space(2)
        
                'Valor do Documento
                aDoc(X).Valor = FormataValor(rsDocumentos!Valor & "", 15)
                sLinha = sLinha & FormataValor(rsDocumentos!Valor, 15)
        
                'Frente e Verso
                aDoc(X).Frente = Trim(rsDocumentos!Frente & "")
                aDoc(X).Verso = Trim(rsDocumentos!Verso & "")
        
                'IdDocto
                aDoc(X).IdDocto = rsDocumentos!IdDocto & ""
        
                'Leitura
                aDoc(X).Leitura = rsDocumentos!Leitura & ""
            
                'Ordem
                aDoc(X).Ordem = rsDocumentos!Ordem & ""
                
                If IsNull(rsDocumentos!SaldoDisponivel) Then
                    'Identificador de Retorno de Saldo sobre conta do cheque
                    aDoc(X).RetornoDeSaldo = False
                    'Saldo Disponivel
                    aDoc(X).SaldoDisponivel = 0
                    'Data e Hora da verificação do saldo
                    aDoc(X).DataHoraSaldo = 0
                    'Valor bloqueado da conta
                    aDoc(X).ValorBloqueado = 0
                    'Valor de limite do cheque
                    aDoc(X).LimiteChequeEspecial = 0
                Else
                    'Identificador de Retorno de Saldo sobre conta do cheque
                    aDoc(X).RetornoDeSaldo = True
                    'Saldo Disponivel
                    aDoc(X).SaldoDisponivel = Format(rsDocumentos!SaldoDisponivel, sMasc)
                    'Data e Hora da verificação do saldo
                    aDoc(X).DataHoraSaldo = Format(rsDocumentos!DataHoraSaldo, "dd/mm/yyyy     hh:mm")
                    'Valor bloqueado da conta
                    aDoc(X).ValorBloqueado = Format(rsDocumentos!ValorBloqueado, sMasc)
                    'Valor de limite do cheque
                    aDoc(X).LimiteChequeEspecial = Format(rsDocumentos!LimiteChequeEspecial, sMasc)
                End If
                'Obtem Agência e Conta qdo. vínculo somente com depósito e existencia de Ajuste
                If IsNull(rsDocumentos!Agencia) And IsNull(rsDocumentos!AgenciaCredito) Then
                    aDoc(X).DepositoAgencia = 0
                    aDoc(X).DepositoConta = 0
                Else
                    aDoc(X).DepositoAgencia = IIf(IsNull(rsDocumentos!AgenciaCredito), rsDocumentos!Agencia, rsDocumentos!AgenciaCredito)
                    aDoc(X).DepositoConta = IIf(IsNull(rsDocumentos!ContaCredito), rsDocumentos!Conta, rsDocumentos!ContaCredito)
                End If
                
                LstDocto.AddItem sLinha
                LstDocto.ItemData(LstDocto.NewIndex) = rsDocumentos!IdDocto
                rsDocumentos.MoveNext
                X = X + 1
                DoEvents
            End If
        Wend
    
        'Verifica se existe diferença na somatória da capa
'        If lstCapa.Enabled Then
            If CapaComDiferenca(rsDocumentos) Then
                AlterouDocto = True
                
                CmdEnviaTransmissao.Enabled = True
                cmdEnviarExpedicao.Enabled = True
            End If
'        End If
    
    Else
        Call HDObjetosImagem(False)
    End If

    LstDocto.Visible = True
    cmdRetiraOcorrencia.Enabled = False
    CmdEnviaCompensacao.Enabled = False
    cmdRemoverVinculo.Enabled = False
    cmdSaldo = False
    
    If lstCapa.ListCount > 0 And LstDocto.ListCount > 0 Then
        LstDocto.Selected(Val(Indice)) = True
        IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa
    End If
    
    'Verifica se existe documento sem vínculo
    If ExisteDoctoSemVinculo Then
        cmdVincular.Enabled = True
    Else
        cmdVincular.Enabled = False
    End If
    
    'Verifica se existe documento sem Estorno
'''    If ExisteDoctoSemEstorno Then
'''        cmdEstorno.Enabled = True
'''    Else
'''        cmdEstorno.Enabled = False
'''    End If
    
    If LstDocto.Visible = True Then
        LstDocto.SetFocus
    End If
    DoEvents
  
    Exit Sub

ERRO_PREENCHELISTDOCTO:
  
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Preencher Lista de Documentos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
  
End Sub

Private Sub cmd_Sld_Fechar_Click()

    lbl_Sld_ValorCheque(1).Caption = "":        lbl_Sld_ValorCheque(1).ForeColor = vbBlack
    lbl_Sld_DataSaldo(1).Caption = "":          lbl_Sld_DataSaldo(1).ForeColor = vbBlack
    lbl_Sld_LimiteCheque(1).Caption = "":       lbl_Sld_LimiteCheque(1).ForeColor = vbBlack
    lbl_Sld_SaldoDisponivel(1).Caption = "":    lbl_Sld_SaldoDisponivel(1).ForeColor = vbBlack
    lbl_Sld_ValorBloqueado(1).Caption = "":     lbl_Sld_ValorBloqueado(1).ForeColor = vbBlack

    Call TelaConsultaSaldo(False)
    
End Sub

Private Sub CmdAtualizar_Click()
    
    If FrmPesquisa.Visible Then Exit Sub
    If frmLocalizar.Visible = True Then Exit Sub
    
    If AlterouDocto Then
        If MsgBox("Capa com diferença de valores, continua com a mesma capa ?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
            Exit Sub
        End If
    End If
    
    If Screen.MousePointer = vbDefault Then
        Screen.MousePointer = vbHourglass

        If AlterouDocto Then
          'A Capa anterior sofreu alteração
            If IdSelecionado <> 0 Then
                Call AtualizaStatusCapa(IdSelecionado, "N")
            End If
        ElseIf IdSelecionado <> 0 Then
            'A Capa anterior não sofreu alteração , Voltar o Status para 'N'
            Call AtualizaStatusCapa(IdSelecionado, "N")
            IdSelecionado = 0
        End If

        Screen.MousePointer = vbDefault

        If Not PreencheListCapas Then
            MsgBox "Não Existem Envelopes / Malotes para C.S.P.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
        End If
    Else
        Call HDMalote(False)
    End If
  
End Sub

Private Sub CmdCancelar_Click()

   frmLocalizar.Visible = False
   LstDocto.SetFocus
   CmdLocalizar.Enabled = True
   
End Sub

Private Sub CmdEnviaCompensacao_Click()

Dim X As Integer
Dim Inicio As Integer, Fim As Integer

On Error GoTo Err_CmdEnviaCompensacao

    If frmLocalizar.Visible = True Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If LstDocto.ListCount = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido alterar documento(s) de capa duplicada.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount > 1 Then
        Inicio = 1
        Fim = LstDocto.ListCount
    Else
        Inicio = LstDocto.ListIndex + 1
        Fim = LstDocto.ListIndex + 1
    End If
    
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            
            'Verificar se documento transmitido
            If aDoc(X).Status = "T" Then
                MsgBox "Documento já transmitido, ação inválida.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento com ocorrência
            If aDoc(X).Status = "F" Or aDoc(X).Status = "D" Or aDoc(X).Status = "C" Then
                MsgBox "Somente poderá ser enviado para compensação documento sem ocorrência.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento com duplicidade
            If aDoc(X).Duplicidade Then
                MsgBox "Somente poderá ser enviado para compensação documento sem duplicidade.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento para Estorno
            If aDoc(X).EstornoDocto Then
                MsgBox "Somente poderá ser enviado para compensação documento sem Estorno.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se o documento é uma Capa
            If Not (aDoc(X).TipoDocto = "5" Or aDoc(X).TipoDocto = "7") Then
                MsgBox "Documento inválido para enviá-lo para compensação", vbInformation, App.Title
                Exit Sub
            End If

        End If
        DoEvents
    Next X

    ' Atualizar Documentos selecionados
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            'Altera tipodocto do documento Cheque de (5)Ubb Sacado para (6)Compensação
            Set qryEnviarCompensacao = Geral.Banco.CreateQuery("", "{? = call EnviaChequeCompensacao (?,?)}")
            With qryEnviarCompensacao
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
                .rdoParameters(2) = aDoc(X).IdDocto       'IdDocto
                .Execute
            End With

            If qryEnviarCompensacao(0).Value = 1 Then
                MsgBox "Ocorreu um erro ao enviar cheque para compensação.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            If Not AtualizaStatusDocumento(aDoc(X).IdDocto, "1") Then
                MsgBox "Ocorreu um erro ao atualizar o status do documento.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            aDoc(X).TipoDocto = 6
            aDoc(X).Status = 1
            aDoc(X).Ocorrencia = ""
        
            Call SetAlteraDocto(True)
            
            'Gravar Log
            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 262)
        End If
    Next X
    
    Call SetAlteraDocto(True)

    Call PreencheListDocto(LstDocto.ListIndex)

    LstDocto.SetFocus
    Exit Sub

Err_CmdEnviaCompensacao:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Retirada da Ocorrência.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub

Private Sub cmdEnviarExpedicao_Click()

Dim X As Integer

    If frmLocalizar.Visible = True Then Exit Sub
    
    'Verificar se existe alguma capa selecionada para Encerramento
    If lstCapa.ListIndex = -1 Then
        MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verifica se existe documento com status <> (T,D,F e C) onde somente poderá ser enviado
    'para expedição os documentos transmitidos e com ocorrência
    If ExisteDoctoPendenteNaoTransmitido Then
        MsgBox "Ação inválida, existe(m) documento(s) não transmitido(s) e sem ocorrência ! ", vbInformation, App.Title
        Exit Sub
    End If

    'Muda status de documento de (F) para (C) - Altera status de Ocorrência gerado pelo robô
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).Status = "F" Then
            If Not AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "C") Then
                MsgBox "Não foi possível atualizar situação do documento" & _
                vbCrLf & vbCrLf & "Status de 'F' para 'C'"
                Exit Sub
            End If
        End If
    Next

    If MsgBox("Enviar Capa para Expedição ?", vbYesNo + vbDefaultButton1) = vbNo Then
        Exit Sub
    End If

    'Excluir Ajustes , se existir
    If Not ExcluiAjuste Then Exit Sub

    'Gera controle de Capa
    Call AtualizaCtrlCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "Capa enviada para expedição.", 31)
    
    'Envia capa para Expedição
    GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 278
    
    'Envia Capa para Expedição
    Call FinalizaCapa("T")
    
    AlterouDocto = False
'    lstCapa.Enabled = True
    
    Screen.MousePointer = vbDefault
    
    If lstCapa.ListIndex = -1 Then
        CmdAtualizar.Enabled = False
        CmdLocalizar.Enabled = False
        CmdOcorrencia.Enabled = False
        cmdRetiraOcorrencia.Enabled = False
        cmdVincular.Enabled = False
'''        cmdEstorno.Enabled = False
        CmdEnviaCompensacao.Enabled = False
        CmdEnviaTransmissao.Enabled = False
        cmdEnviarExpedicao.Enabled = False
        cmdRemoverVinculo.Enabled = False
        cmdSaldo.Enabled = False
        cmdFechar.SetFocus
    Else
        If (LstDocto.Visible = True) And (LstDocto.Enabled = True) Then LstDocto.SetFocus
    End If
   
    Exit Sub

End Sub

Private Sub CmdEnviaTransmissao_Click()

    Dim X As Integer
    Dim Vinculo As Long
    Dim cTotalCred As Currency, cTotalDeb As Currency
    Dim cDifCred As Currency, cDifDeb As Currency
    Dim iAgenciaAjuste As Integer, lContaAjuste As Long
    Dim bSomenteDeposito As Boolean
    Dim bAgenciaContaDeposito As Boolean
    Dim avinculo()  As Long
    Dim iVinculo As Integer
    Dim bVerificado As Boolean

    Dim lUnicoVinculo As Long
    
    ReDim avinculo(0)
    
    tpAjuste.Agencia = 0
    tpAjuste.Conta = 0
    tpAjuste.TipoDocto = 0
    tpAjuste.Valor = 0
    tpAjuste.Vinculo = 0
    
    If frmLocalizar.Visible = True Then Exit Sub
    
    'Verificar se existe alguma capa selecionada para Encerramento
    If lstCapa.ListIndex = -1 Then
        MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
        Exit Sub
    End If
    
    If ExisteDoctoSemVinculo Then
        MsgBox "Não é permitido enviar para transmissão capa com documento(s) sem vínculo.", vbInformation, App.Title
        Exit Sub
    End If
    
    lUnicoVinculo = 0
    lUnicoVinculo = ExisteDoctoUnicoVinculo()
    If lUnicoVinculo <> 0 Then
        MsgBox "Não existe referência para documento com vínculo " & CStr(lUnicoVinculo) & " , favor verificar !", vbInformation, App.Title
        Exit Sub
    End If
    
    'Verifica se para cada vínculo existe Contra Partida para Débito e Credito
    If Not ExisteContraPartidaPorVinculo() Then Exit Sub
    
    If MsgBox("Enviar Capa para Transmissão ?", vbYesNo + vbDefaultButton1) = vbNo Then
        Exit Sub
    End If
    
    If Not AcertaTipoDocto() Then Exit Sub
    
    'Excluir Ajustes , se existir
    If Not ExcluiAjuste Then Exit Sub

    'Fazer a conferência de valores vínculo à vínculo
    Vinculo = 0
    Screen.MousePointer = vbHourglass
    
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).TipoDocto <> 1 And aDoc(X + 1).Vinculo <> 0 Then

            If Vinculo <> aDoc(X + 1).Vinculo Then
                Vinculo = aDoc(X + 1).Vinculo
                
                'Verifica se vinculo já sofreu verificação
                bVerificado = False
                For iVinculo = 1 To UBound(avinculo)
                    If avinculo(iVinculo) = Vinculo Then
                        bVerificado = True
                        Exit For
                    End If
                Next
                
                If Not bVerificado Then
                    ReDim Preserve avinculo(UBound(avinculo) + 1)
                    avinculo(UBound(avinculo)) = Vinculo
                End If
                
                If Not bVerificado Then
                                    
                    cTotalCred = 0
                    cTotalDeb = 0
                    bSomenteDeposito = False
                    bAgenciaContaDeposito = False
                    
                    Call SomaDebitosCreditos(aDoc(X + 1).Vinculo, cTotalCred, cTotalDeb, bSomenteDeposito)
                    
                    If cTotalCred <> cTotalDeb Then
                        'Se Vínculo somente com Depósito/OCT, verificar se todos chq's são tipo (7)Depósito
                        If bSomenteDeposito Then
                            'Se Todos chq's depósito, obter a agência e conta do Depósito/OCT
                            Call AgenciaContaSomenteDeposito(aDoc(X + 1).Vinculo, iAgenciaAjuste, lContaAjuste)
                            If iAgenciaAjuste <> 0 Then bAgenciaContaDeposito = True
                        End If
                        
                        'Verifica se já obteve agência e Conta para Ajustes
                        If iAgenciaAjuste = 0 Then
                            Do While True
                                Screen.MousePointer = vbDefault
                                
                                cDifCred = 0
                                cDifDeb = 0
                                Call DiferencaParaAjuste(aDoc(X + 1).Vinculo, cDifCred, cDifDeb)
                                If Not DigitaAgenciaConta(aDoc(X + 1).Vinculo, iAgenciaAjuste, lContaAjuste, (cDifCred - cDifDeb)) Then
                                    If MsgBox("Existe diferença de valores, favor digitar Agência e Conta para Ajuste", vbOKOnly + vbOKCancel + vbDefaultButton1, App.Title) = vbCancel Then
                                        GoTo err_saida
                                    End If
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If
                        Screen.MousePointer = vbHourglass
                        
                        'Gera ajuste para os documentos referentes ao vínculo
                        tpAjuste.Agencia = iAgenciaAjuste
                        tpAjuste.Conta = lContaAjuste
                        'Se vínculo somente com depósito gera somente ajuste (não automático)
                        If bSomenteDeposito Then
                            tpAjuste.TipoDocto = IIf((cTotalCred - cTotalDeb) > 0, 32, 33)
                        Else
                            tpAjuste.TipoDocto = IIf((cTotalCred - cTotalDeb) > 0, 34, 38)
                        End If
                        tpAjuste.Valor = Abs(cTotalCred - cTotalDeb)
                        tpAjuste.Vinculo = Vinculo
                        
                        If Not GeraAjuste() Then GoTo err_saida
                        
                        'Para Agencia e Conta de Depósito/OCT, utilizar somente no mesmo Vínculo
                        If bAgenciaContaDeposito Then
                            iAgenciaAjuste = 0
                            lContaAjuste = 0
                            tpAjuste.Agencia = 0
                            tpAjuste.Conta = 0
                        End If
                        
                        'Grava log para ajuste
                        GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 271

                    End If
                End If
            End If
        End If
    Next
    
    'Muda status de documento de (F) para (C) - Altera status de Ocorrência gerado pelo robô
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).Status = "F" Then
            If Not AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "C") Then
                MsgBox "Não foi possível atualizar situação do documento" & _
                vbCrLf & vbCrLf & "Status de 'F' para 'C'"
                Exit Sub
            End If
        End If
    Next
    
    'Gera controle de Capa
    Call AtualizaCtrlCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "Capa enviada para transmissão.", 31)
    
    'Envia capa para transmissão
    GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 266
    
    'Envia Capa para transmissâo
    Call FinalizaCapa("R")
    
    AlterouDocto = False
'    lstCapa.Enabled = True
    
    Screen.MousePointer = vbDefault
    
    If lstCapa.ListIndex = -1 Then
        CmdAtualizar.Enabled = False
        CmdLocalizar.Enabled = False
        CmdOcorrencia.Enabled = False
        cmdRetiraOcorrencia.Enabled = False
        cmdVincular.Enabled = False
'''        cmdEstorno.Enabled = False
        CmdEnviaCompensacao.Enabled = False
        CmdEnviaTransmissao.Enabled = False
        cmdEnviarExpedicao.Enabled = False
        cmdRemoverVinculo.Enabled = False
        cmdSaldo.Enabled = False
        cmdFechar.SetFocus
    Else
        If (LstDocto.Visible = True) And (LstDocto.Enabled = True) Then LstDocto.SetFocus
    End If
   
   
Exit Sub

err_saida:
    Screen.MousePointer = vbDefault
    Call CmdAtualizar_Click
    DoEvents

End Sub

Private Sub cmdEstorno_Click()

    Estorno.m_lngIdCapaCSP = aCapa(lstCapa.ListIndex + 1).IdCapa
    Estorno.m_strNumCapa = aCapa(lstCapa.ListIndex + 1).Capa
    Estorno.m_IndexDoc = IIf(LstDocto.ListIndex <= 0, 1, aDoc(LstDocto.ListIndex + 1).NrSeq - 1)
    Estorno.Show vbModal, Me
    
    If Estorno.m_bHouveEstornoCSP Then
        Call PreencheListDocto(LstDocto.ListIndex)
        
        Call SetAlteraDocto(True)
    
        LstDocto.SetFocus
        lstDocto_Click
    End If
    
End Sub

Private Sub CmdFechar_Click()

    If AlterouDocto Then
'        MsgBox "Foi alterado documento(s) nesta capa, favor verificar e encerrar a capa.", vbInformation, App.Title
        If MsgBox("Capa com diferença de valores, deseja realmente fechar ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub
Private Sub CmdFecharPesquisa_Click()

  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  'Call GravaLog(0, 0, 268)
  
  Call CmdFechar_Click
  
End Sub

Public Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
  'poi, o canon não gera verso.
  If (aDoc(LstDocto.ListIndex + 1).Ordem = "0") Or (aDoc(LstDocto.ListIndex + 1).Ordem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & aDoc(LstDocto.ListIndex + 1).Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(LstDocto.ListIndex + 1).Frente, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(LstDocto.ListIndex + 1).Ordem = "2") Then
               .Intensity 140
            Else
               .Intensity 220
            End If
           .PaintZoomFactor = 100
           .AutoRepaint = True
        End With
    Else
        Lead1.Tag = "V"     'se frente, mostrar verso
        With Lead1
            .AutoRepaint = False
  
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Trim(aDoc(LstDocto.ListIndex + 1).Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(LstDocto.ListIndex + 1).Verso, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(LstDocto.ListIndex + 1).Ordem = "2") Then
               .Intensity 140
            Else
               .Intensity 220
            End If
            .PaintZoomFactor = 100
            .AutoRepaint = True
        End With
    End If
  End If
  
  DoEvents
  teclou = False
  Exit Sub

ERRO_FRENTEVERSO:
  FrmImagem.Visible = False
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Public Sub cmdInverteCor_Click()

  On Error GoTo ERRO_INVERTECOR

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  Lead1.Invert
  DoEvents
  teclou = False
  Exit Sub

ERRO_INVERTECOR:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub CmdLocalizar_Click()

   If FrmPesquisa.Visible = False Then
      frmLocalizar.Visible = True
      txtNumEnvMal.SetFocus
      CmdLocalizar.Enabled = False
   End If
   
End Sub
Private Sub cmdOcorrencia_Click()

Dim X As Integer
Dim Valor As Currency
Dim Inicio As Integer, Fim As Integer
Dim IdDocto As Long
Dim strDescricao As String

On Error GoTo Err_cmdOcorrencia

    If frmLocalizar.Visible Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If LstDocto.ListCount = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido gerar ocorrência para documento(s) de capa duplicada.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount > 1 Then
        Inicio = 1
        Fim = LstDocto.ListCount
    Else
        Inicio = LstDocto.ListIndex + 1
        Fim = LstDocto.ListIndex + 1
    End If
    
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            
            'Verifica se Documento transmitido
            If aDoc(X).Status = "T" Then
                MsgBox "Não é permitido gerar Ocorrência para documento já transmitido", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verifica se Documento com ocorrência gerada pelo MDI
            If aDoc(X).Status = "D" Then
                MsgBox "Documento com ocorrência gerada pelo sistema MDI, operação negada.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verifica se Documento é Capa
            If aDoc(X).TipoDocto = "1" Then
                MsgBox "Não é permitido gerar Ocorrência para Capa de " & LblEnv_Mal.Caption & ".", vbInformation, App.Title
                Exit Sub
            End If

            'Verifica se Documento duplicado
            If aDoc(X).Duplicidade Then
                MsgBox "Não é permitido gerar Ocorrência para Documento(s) Duplicado(s).", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento para Estorno
            If aDoc(X).EstornoDocto Then
                MsgBox "Não é permitido gerar Ocorrência de documento para Estorno.", vbInformation, App.Title
                Exit Sub
            End If
            'Guarda Um IdDocto para consulta de complemento da ocorrência
            IdDocto = aDoc(X).IdDocto
        End If
        DoEvents
    Next X

    'Abre tela de ocorrência passando parâmetro para não apresentar o botão
    'de remoção de ocorrência
    Ocorrencia.cmdRemoverOcorrencia.Tag = "CSP"
    
    'Busca descrição do complemento de ocorrência, caso exista
    strDescricao = ""
'''    Call GravaComplementoOcorrencia(IdDocto, "C", strDescricao)
    
    Ocorrencia.m_Descricao = Trim(strDescricao)
    
    Ocorrencia.Show vbModal, Me

    If Ocorrencia.Result Then
        Call SetAlteraDocto(True)

        'Foi escolhida uma Ocorrência -> Atualizar Documentos selecionados
        For X = 0 To LstDocto.ListCount - 1
            If LstDocto.Selected(X) = True Then
                'Atualizar o Campo 'OCORRENCIA'
                Set qryAtualizaOcorrencia = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrenciaDocumento (?,?,?)}")
                With qryAtualizaOcorrencia
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
                    .rdoParameters(2) = aDoc(X + 1).IdDocto      'IdDocto
                    .rdoParameters(3) = Ocorrencia.CodOcorr      'Código da Ocorrencia
                    .Execute
                End With

                If qryAtualizaOcorrencia(0).Value = 1 Then
                    MsgBox "Ocorreu um erro ao atualizar a ocorrência do documento.", vbInformation + vbOKOnly, App.Title
                    Exit Sub
                End If
                
                'Grava/Altera ou Exclui Complemento da Ocorrência
'''                If Not GravaComplementoOcorrencia(aDoc(X + 1).IdDocto, IIf(Ocorrencia.m_Descricao = "", "E", "G"), Ocorrencia.m_Descricao) Then Exit Sub

                If Not AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "F") Then
                    MsgBox "Ocorreu um erro ao atualizar o status do documento.", vbInformation + vbOKOnly, App.Title
                    Exit Sub
                End If

                aDoc(X + 1).Status = "F"
                aDoc(X + 1).Ocorrencia = Ocorrencia.CodOcorr
                
                'Gravar Log
                Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X + 1).IdDocto, 264)
            End If
        Next X

        Unload Ocorrencia
        Call PreencheListDocto(LstDocto.ListIndex)
    
        Call SetAlteraDocto(True)
    
    Else
        Unload Ocorrencia
    End If

    LstDocto.SetFocus
    lstDocto_Click
    Exit Sub

Err_cmdOcorrencia:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Ocorrência.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub
Private Sub cmdProcurar_Click()

    Dim iIndex                   As Integer
    Dim Encontrou                As Boolean
    Dim qryGetDescStatusCapa     As rdoQuery 'Pega descricao do status da capa
    Dim sCapa                    As String
    
    
    If (Trim(txtNumEnvMal.Text) = "") Or (Not IsNumeric(txtNumEnvMal.Text)) Then
        MsgBox "Entre com um número de capa válido.", vbExclamation
        txtNumEnvMal.SetFocus
        Exit Sub
    End If
    
    Set qryGetDescStatusCapa = Geral.Banco.CreateQuery("", "{Call GetDescStatusCapa(?,?,?)}")

    If Trim(txtNumEnvMal.Text) <> "" Then
        If IsNumeric(txtNumEnvMal.Text) Then
            'Atualizar a lista de capas antes da pesquisa
            Call CmdAtualizar_Click
            
            'Verificar se a capa informada está na lista de capas
            For iIndex = 0 To lstCapa.ListCount - 1
                If CDbl(lstCapa.List(iIndex)) = CDbl(txtNumEnvMal.Text) Then
                    lstCapa.Selected(iIndex) = True
                    Encontrou = True
                    Exit For
                End If
                DoEvents
            Next iIndex
        End If
    End If

    sCapa = txtNumEnvMal.Text
    txtNumEnvMal.Text = ""
    frmLocalizar.Visible = False
    CmdLocalizar.Enabled = True
    
    'Verificar se encontrou a capa
    If Not Encontrou Then
    
        With qryGetDescStatusCapa
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = CDbl(sCapa)
            .rdoParameters(2).Direction = rdParamOutput
            .Execute
            
            If Trim(.rdoParameters(2).Value) <> "" Then
                MsgBox .rdoParameters(2).Value, vbInformation
            Else
                MsgBox "Capa não Encontrada.", vbInformation + vbOKOnly, App.Title
            End If
            
        End With
    
        LblEnv_Mal.Caption = ""
        
        Call HDMalote(False)
        lblNumMalote.Caption = ""
        
        If IdSelecionado <> 0 Then
            'A Capa anterior não sofreu alteração , Voltar o Status para 'N'
            Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "N")
        End If
        
        lstCapa.ListIndex = -1
        LstDocto.Clear
        FrmImagem.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdRemoverVinculo_Click()

Dim X As Integer
Dim Inicio As Integer, Fim As Integer

On Error GoTo Err_cmdRemoverVinculo

    If frmLocalizar.Visible = True Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If LstDocto.ListCount = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido remover vínculo de capa duplicada.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount > 1 Then
        Inicio = 1
        Fim = LstDocto.ListCount
    Else
        Inicio = LstDocto.ListIndex + 1
        Fim = LstDocto.ListIndex + 1
    End If
    
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            
            'Verificar se documento transmitido
            If aDoc(X).Status = "T" Then
                MsgBox "Documento já transmitido, ação inválida.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento com ocorrência
            If aDoc(X).Status = "F" Or aDoc(X).Status = "D" Or aDoc(X).Status = "C" Then
                MsgBox "Somente poderá ser removido vínculo para documento sem ocorrência.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento com duplicidade
            If aDoc(X).Duplicidade Then
                MsgBox "Somente poderá ser removido vínculo para documento sem duplicidade.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento para Estorno
            If aDoc(X).EstornoDocto Then
                MsgBox "Somente poderá ser removido vínculo para documento sem Estorno.", vbInformation, App.Title
                Exit Sub
            End If

        End If
        DoEvents
    Next X

    ' Atualizar Documentos selecionados
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) And aDoc(X).Vinculo <> 0 Then
            'Altera tipodocto do documento Cheque de (5)Ubb Sacado para (6)Compensação
            Set qryGeraVinculo = Geral.Banco.CreateQuery("", "{? = call GeraVinculoDocumento (?,?,?)}")
            With qryGeraVinculo
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                .rdoParameters(2) = aDoc(X).IdDocto             'IdDocto do documento
                .rdoParameters(3) = 0                           'IdDocto como vínculo
                .Execute
            End With

            If qryGeraVinculo(0).Value = 1 Then
                MsgBox "Ocorreu um erro ao vincular documento.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            aDoc(X).Vinculo = 0
        
            Call SetAlteraDocto(True)
            
        End If
    Next X
    
    Call SetAlteraDocto(True)

    Call PreencheListDocto(LstDocto.ListIndex)

    LstDocto.SetFocus
    Exit Sub

Err_cmdRemoverVinculo:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Retirada de Vínculo.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub

Private Sub cmdRetiraOcorrencia_Click()

Dim X As Integer
Dim Inicio As Integer, Fim As Integer

On Error GoTo Err_cmdRetiraOcorrencia

    If frmLocalizar.Visible = True Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If LstDocto.ListCount = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido retirar ocorrência para documento(s) de capa duplicada.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount > 1 Then
        Inicio = 1
        Fim = LstDocto.ListCount
    Else
        Inicio = LstDocto.ListIndex + 1
        Fim = LstDocto.ListIndex + 1
    End If
    
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) = True Then
            
            'Verifica se Documento transmitido
            If aDoc(X).Status = "T" Then
                MsgBox "Não é permitido retirar Ocorrência para documento já transmitido", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verifica se Documento com ocorrência gerada pelo MDI
            If aDoc(X).Status = "D" Then
                MsgBox "Documento com ocorrência gerada pelo sistema MDI, operação negada.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verifica se Documento sem marcação de ocorrência = "F"
            If aDoc(X).Status <> "F" And aDoc(X).Status <> "C" Then
                MsgBox "Somente será permitido retirar ocorrência para documento com ocorrência.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verifica se Documento é Capa
            If aDoc(X).TipoDocto = "1" Then
                MsgBox "Não é permitido retirar Ocorrência para Capa de " & LblEnv_Mal.Caption & ".", vbInformation, App.Title
                Exit Sub
            End If

            'Verificar se o documento duplicado
            If aDoc(X).Duplicidade = True Then
                MsgBox "Não é permitido retirar Ocorrência para Documento(s) Duplicado(s).", vbInformation, App.Title
                Exit Sub
            End If

        End If
        DoEvents
    Next X

    'Foi escolhida uma Ocorrência -> Atualizar Documentos selecionados
    For X = 0 To LstDocto.ListCount - 1
        If LstDocto.Selected(X) = True Then
            'Atualizar o Campo 'OCORRENCIA'
            Set qryAtualizaOcorrencia = Geral.Banco.CreateQuery("", "{? = call RetirarOcorrenciaDocumento (?,?)}")
            With qryAtualizaOcorrencia
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
                .rdoParameters(2) = aDoc(X + 1).IdDocto      'IdDocto
                .Execute
            End With

            If qryAtualizaOcorrencia(0).Value = 1 Then
                MsgBox "Ocorreu um erro ao atualizar a retirada de ocorrência do documento.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            'Exclui Complemento da Ocorrência
'''            If Not GravaComplementoOcorrencia(aDoc(X + 1).IdDocto, "E", "") Then Exit Sub

            If Not AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "1") Then
                MsgBox "Ocorreu um erro ao atualizar o status do documento.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            aDoc(X + 1).Status = "1"
            aDoc(X + 1).Ocorrencia = ""

            'Gravar Log
            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X + 1).IdDocto, 263)
        End If
    Next X

    Call SetAlteraDocto(True)
    
    Call PreencheListDocto(LstDocto.ListIndex)

    LstDocto.SetFocus
    Exit Sub

Err_cmdRetiraOcorrencia:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Retirada da Ocorrência.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub

Public Sub cmdRotacao_Click()

  On Error GoTo ERRO_ROTACAO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  Lead1.FastRotate 90
  DoEvents
  teclou = False
  Exit Sub

ERRO_ROTACAO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub cmdSaldo_Click()
    
Dim X As Integer

    If LstDocto.SelCount < 1 Then Exit Sub
    
    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount <> 1 Then
        Beep
        MsgBox "Para consulta de Saldo, favor selecionar somente um cheque." & vbCrLf & vbCrLf & _
                "* Somente serão consultados cheques do Unibanco", vbInformation + vbOKOnly, App.Title
        Exit Sub
    Else
        X = LstDocto.ListIndex + 1
    End If
    
    'Verifica se docto selecionado é Cheque UBB
    If Not (InStr("5,6", aDoc(LstDocto.ListIndex + 1).TipoDocto) <> 0 And InStr("409,230", Left(aDoc(LstDocto.ListIndex + 1).Leitura, 3)) <> 0) Then
        MsgBox "Para consulta de Saldo, favor selecionar somente cheque do Unibanco.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    'Verificar se documento transmitido
    If aDoc(X).Status = "T" Then
        MsgBox "Documento transmitido não poderá ser consultado.", vbInformation, App.Title
        Exit Sub
    End If
        
    'Verifica se existe retorno de saldo
    If Not aDoc(X).RetornoDeSaldo Then
        MsgBox "Consulta de Saldo indisponível para este cheque.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Apresenta tela com consulta de saldo
    Call TelaConsultaSaldo(True)
    
    fraSaldo.Left = ((Screen.Width - fraSaldo.Width) / 2) - 800
    fraSaldo.Top = 700
    
    lbl_Sld_ValorCheque(1).Caption = Format(aDoc(X).Valor, sMasc)
    
    lbl_Sld_DataSaldo(1).Caption = aDoc(X).DataHoraSaldo
    
    lbl_Sld_LimiteCheque(1).Caption = aDoc(X).LimiteChequeEspecial
    If aDoc(X).LimiteChequeEspecial < 0 Then lbl_Sld_LimiteCheque(1).ForeColor = vbRed
    
    lbl_Sld_SaldoDisponivel(1).Caption = aDoc(X).SaldoDisponivel
    If aDoc(X).SaldoDisponivel < 0 Then lbl_Sld_SaldoDisponivel(1).ForeColor = vbRed
    
    lbl_Sld_ValorBloqueado(1).Caption = aDoc(X).ValorBloqueado
    If aDoc(X).ValorBloqueado < 0 Then lbl_Sld_ValorBloqueado(1).ForeColor = vbRed

End Sub

Private Sub cmdVincular_Click()

Dim X As Integer
Dim Inicio As Integer, Fim As Integer

'Acumuladores para soma de documento do mesmo tipo
Dim iDoctoDeposito As Integer, iDoctoOCT As Integer, iDoctoLanctoInterno As Integer
Dim iDoctoADCC As Integer, iDoctoCheque As Integer, iDoctoPagto As Integer
Dim iChqUBB As Integer, iChqOutros As Integer
Dim cTotalCred As Currency, cTotalDeb As Currency
Dim cTotDebitos As Currency, cTotCreditos As Currency

'Acumulador do IdDocto do Primeiro documento do mesmo tipo
Dim lIdDoctoDeposito As Long, lIdDoctoOCT As Long, lIdDoctoLanctoInterno As Long
Dim lIdDoctoADCC As Long, lIdDoctoCheque As Long, lIdDoctoPagto As Long
Dim lIdDoctoVincular As Long

'Dados de documento transmitido
Dim iDoctoTransm As Integer, lIdDoctoTransm As Long

On Error GoTo err_cmdVincular_Click

    If frmLocalizar.Visible = True Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If LstDocto.ListCount = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido alterar documento(s) de capa duplicada.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If LstDocto.SelCount > 1 Then
        Inicio = 1
        Fim = LstDocto.ListCount
    Else
        Inicio = LstDocto.ListIndex + 1
        Fim = LstDocto.ListIndex + 1
    End If
    
    iDoctoDeposito = 0:     iDoctoOCT = 0:      iDoctoLanctoInterno = 0
    iDoctoADCC = 0:         iDoctoCheque = 0:   iDoctoPagto = 0
    iChqUBB = 0:            iChqOutros = 0
    
    lIdDoctoDeposito = 0:   lIdDoctoOCT = 0:    lIdDoctoLanctoInterno = 0
    lIdDoctoADCC = 0:       lIdDoctoCheque = 0: lIdDoctoPagto = 0:      lIdDoctoVincular = 0
    
    iDoctoTransm = 0:       lIdDoctoTransm = 0
    cTotDebitos = 0:        cTotCreditos = 0
    
    '-----------------------------------------------------------------------------------------
    '       Soma-se quantidade de documentos Apenas Selecionados independentes de vínculo
    '-----------------------------------------------------------------------------------------
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            
            'Verificar se documento transmitido
            If aDoc(X).Status = "T" Then
                If lIdDoctoTransm <> aDoc(X).Vinculo Then
                    iDoctoTransm = iDoctoTransm + 1
                    lIdDoctoTransm = aDoc(X).Vinculo
                End If
            End If
            
            'Verificar se documento com ocorrência
            If aDoc(X).Status = "F" Or aDoc(X).Status = "D" Or aDoc(X).Status = "C" Then
                MsgBox "Documento com ocorrência não poderá ser vinculado.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento com duplicidade
            If aDoc(X).Duplicidade Then
                MsgBox "Documento com duplicidade não poderá ser vinculado.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se documento Para Estorno
            If aDoc(X).EstornoDocto Then
                MsgBox "Documento para Estorno não poderá ser vinculado.", vbInformation, App.Title
                Exit Sub
            End If
            
            'Verificar se o documento é uma Capa
            If aDoc(X).TipoDocto = "1" Then
                MsgBox "Documento Capa não poderá ser vinculado.", vbInformation, App.Title
                Exit Sub
            End If
            '---------------------------------------------------------------------------
            '               Soma quantidade por tipo de documento
            '---------------------------------------------------------------------------
            Select Case aDoc(X).TipoDocto
                Case 2, 3           ' Depositos
                    iDoctoDeposito = iDoctoDeposito + 1
                    cTotDebitos = cTotDebitos + aDoc(X).Valor
                    If iDoctoDeposito = 1 Then lIdDoctoDeposito = aDoc(X).IdDocto
                Case 4              ' ADCC
                    iDoctoADCC = iDoctoADCC + 1
                    cTotCreditos = cTotCreditos + aDoc(X).Valor
                    If iDoctoADCC = 1 Then lIdDoctoADCC = aDoc(X).IdDocto
                Case 5, 6, 7        ' Cheques
                    iDoctoCheque = iDoctoCheque + 1
                    cTotCreditos = cTotCreditos + aDoc(X).Valor
                    If iDoctoCheque = 1 Then lIdDoctoCheque = aDoc(X).IdDocto
                    'Verifica se cheque UBB
                    If InStr("409*230", Left(aDoc(X).Leitura, 3)) <> 0 Then
                        iChqUBB = iChqUBB + 1
                    Else
                        iChqOutros = iChqOutros + 1
                    End If
                Case 37             ' OCT
                    iDoctoOCT = iDoctoOCT + 1
                    cTotDebitos = cTotDebitos + aDoc(X).Valor
                    If iDoctoOCT = 1 Then lIdDoctoOCT = aDoc(X).IdDocto
                Case 41             ' Lancamento Interno
                    iDoctoLanctoInterno = iDoctoLanctoInterno + 1
                    cTotCreditos = cTotCreditos + aDoc(X).Valor
                    If iDoctoLanctoInterno = 1 Then lIdDoctoLanctoInterno = aDoc(X).IdDocto
                Case 33, 38, 43, 45     'Ajuste de Crédito
                    cTotCreditos = cTotCreditos + aDoc(X).Valor
                Case 32, 34, 42, 44     'Ajuste de Débito
                    cTotDebitos = cTotDebitos + aDoc(X).Valor
                Case 39                 'Capa de OCT não tratar
                Case Else                ' Doctos para Pagamento
                    iDoctoPagto = iDoctoPagto + 1
                    cTotDebitos = cTotDebitos + aDoc(X).Valor
                    If iDoctoPagto = 1 Then lIdDoctoPagto = aDoc(X).IdDocto
            End Select
        End If
        DoEvents
    Next X

    '---------------------------------------------------------------------------
    '       Verifica se existe documento débito e crédito para vínculo
    '---------------------------------------------------------------------------
    If iDoctoDeposito > 0 And iDoctoLanctoInterno = 0 And iDoctoCheque = 0 And iDoctoADCC = 0 Then
        MsgBox "Depósito sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoCheque > 0 And iDoctoPagto = 0 And iDoctoOCT = 0 And iDoctoDeposito = 0 Then
        MsgBox "Cheque(s) sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoOCT > 0 And iDoctoLanctoInterno = 0 And iDoctoCheque = 0 And iDoctoADCC = 0 Then
        MsgBox "OCT sem cheque para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoADCC > 0 And iDoctoPagto = 0 And iDoctoOCT = 0 And iDoctoDeposito = 0 Then
        MsgBox "ADCC sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
        
    If iDoctoLanctoInterno > 0 And iDoctoPagto = 0 And iDoctoDeposito = 0 And iDoctoOCT = 0 Then
        MsgBox "Lançamento Interno sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoPagto > 0 And iDoctoCheque = 0 And iDoctoLanctoInterno = 0 And iDoctoADCC = 0 Then
        MsgBox "Pagamento sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If

    'Verifica se existe vínculos diferentes para mais de um documento transmitido
    If iDoctoTransm > 1 Then
        MsgBox "Documentos selecionados e transmitidos com vínculos diferentes não poderão ser vinculados.", vbInformation, App.Title
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------------------
    '   Verifica regra para desdobramento de cheque Apenas para documentos selecionados
    '---------------------------------------------------------------------------------------

    If iDoctoPagto > 0 And iDoctoCheque > 0 And iDoctoDeposito > 0 And iChqOutros > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoCheque > 0 And iDoctoDeposito > 1 And iChqOutros > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
'    If iDoctoCheque > 0 And iDoctoLanctoInterno > 0 And iDoctoDeposito > 0 Then
'        MsgBox "Não é permitido o mesmo vínculo para Cheque com desdobramento , favor verificar!", vbInformation, App.Title
'        Exit Sub
'    End If

    If iDoctoCheque > 0 And iDoctoDeposito > 0 And iDoctoADCC > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque com desdobramento , favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoOCT > 0 And (iDoctoPagto > 0 Or iDoctoDeposito > 0 Or iDoctoADCC > 0) Then
        MsgBox "Não é permitido o mesmo vínculo para OCT com PAGTO/ADCC ou DEPÓSITO, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    
    '---------------------------------------------------------------------------
    '   Determina qual IdDocto para vincular todos documentos selecionados
    '---------------------------------------------------------------------------
    If iDoctoTransm = 1 Then
        lIdDoctoVincular = lIdDoctoTransm
    Else
        If iDoctoOCT >= 1 Then                      ' OCT
            lIdDoctoVincular = lIdDoctoOCT
        ElseIf iDoctoLanctoInterno >= 1 Then        ' Lancamento Interno
            lIdDoctoVincular = lIdDoctoLanctoInterno
        ElseIf iDoctoADCC = 1 Then                  ' ADCC
            lIdDoctoVincular = lIdDoctoADCC
        ElseIf iDoctoDeposito = 1 Then              ' Depositos
            lIdDoctoVincular = lIdDoctoDeposito
        ElseIf iDoctoCheque = 1 Then                ' Cheques
            lIdDoctoVincular = lIdDoctoCheque
        ElseIf iDoctoPagto = 1 Then                 ' Doctos para Pagamento
            lIdDoctoVincular = lIdDoctoPagto
        ElseIf iDoctoADCC >= 1 Then                 ' ADCC
            lIdDoctoVincular = lIdDoctoADCC
        ElseIf iDoctoDeposito >= 1 Then             ' Depositos
            lIdDoctoVincular = lIdDoctoDeposito
        ElseIf iDoctoCheque >= 1 Then               ' Cheques
            lIdDoctoVincular = lIdDoctoCheque
        ElseIf iDoctoPagto >= 1 Then                ' Doctos para Pagamento
            lIdDoctoVincular = lIdDoctoPagto
        End If
    End If
    
    '---------------------------------------------------------------------------
    '                   ** COMPLEMENTO NA SOMA DE DOCTOS. **
    '       Soma-se a quantidade de documentos NÃO SELECIONADOS por vínculo,
    '       com isso será somado os documentos com o mesmo nr. de vínculo
    '       contidos fora dos doctos selecionados, com isso obtem-se o total de
    '       doctos pertencentes a um vínculo independente de seleção no ListBox.
    '---------------------------------------------------------------------------
    For X = Inicio To Fim
        If Not (LstDocto.Selected(X - 1)) Then
            If aDoc(X).Vinculo = lIdDoctoVincular Then
                'Verificar se documento com ocorrência
                If Not (aDoc(X).Status = "F" Or aDoc(X).Status = "D" Or aDoc(X).Status = "C") Then
                    'Verificar se documento com duplicidade
                    If Not (aDoc(X).Duplicidade) Then
                        'Verificar se documento Para Estorno
                        If Not (aDoc(X).EstornoDocto) Then
                            'Verificar se o documento é uma Capa
                            If Not (aDoc(X).TipoDocto = "1") Then
            
                                '-------------------------------------------------
                                '       Soma quantidade por tipo de documento
                                '-------------------------------------------------
                                Select Case aDoc(X).TipoDocto
                                    Case 2, 3           ' Depositos
                                        iDoctoDeposito = iDoctoDeposito + 1
                                        cTotDebitos = cTotDebitos + aDoc(X).Valor
                                    Case 4              ' ADCC
                                        iDoctoADCC = iDoctoADCC + 1
                                        cTotCreditos = cTotCreditos + aDoc(X).Valor
                                    Case 5, 6, 7        ' Cheques
                                        iDoctoCheque = iDoctoCheque + 1
                                        cTotCreditos = cTotCreditos + aDoc(X).Valor
                                        'Verifica se cheque UBB
                                        If InStr("409*230", Left(aDoc(X).Leitura, 3)) <> 0 Then
                                            iChqUBB = iChqUBB + 1
                                        Else
                                            iChqOutros = iChqOutros + 1
                                        End If
                                    Case 37             ' OCT
                                        iDoctoOCT = iDoctoOCT + 1
                                        cTotDebitos = cTotDebitos + aDoc(X).Valor
                                    Case 41             ' Lancamento Interno
                                        iDoctoLanctoInterno = iDoctoLanctoInterno + 1
                                        cTotCreditos = cTotCreditos + aDoc(X).Valor
                                    Case 33, 38, 43, 45     'Ajuste de Crédito
                                        cTotCreditos = cTotCreditos + aDoc(X).Valor
                                    Case 32, 34, 42, 44     'Ajuste de Débito
                                        cTotDebitos = cTotDebitos + aDoc(X).Valor
                                    Case 39             'Capa de OCT não tratar
                                    Case Else           ' Doctos para Pagamento
                                        iDoctoPagto = iDoctoPagto + 1
                                        cTotDebitos = cTotDebitos + aDoc(X).Valor
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
        End If
        DoEvents
    Next X
    
    '---------------------------------------------------------------------------------------------------
    '   Verifica regra para desdobramento de cheque Para documentos selecionados e Não selecionados
    '---------------------------------------------------------------------------------------------------
    If iDoctoPagto > 0 And iDoctoCheque > 0 And iDoctoDeposito > 0 And iChqOutros > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoCheque > 0 And iDoctoDeposito > 1 And iChqOutros > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar !", vbInformation, App.Title
        Exit Sub
    End If
    
'    If iDoctoCheque > 0 And iDoctoLanctoInterno > 0 And iDoctoDeposito > 0 Then
'        MsgBox "Não é permitido o mesmo vínculo para Cheque com desdobramento , favor verificar!", vbInformation, App.Title
'        Exit Sub
'    End If
    
    If iDoctoCheque > 0 And iDoctoDeposito > 0 And iDoctoADCC > 0 Then
        MsgBox "Não é permitido o mesmo vínculo para Cheque com desdobramento , favor verificar !", vbInformation, App.Title
        Exit Sub
    End If
    
    If iDoctoOCT > 0 And (iDoctoPagto > 0 Or iDoctoDeposito > 0 Or iDoctoADCC > 0) Then
        MsgBox "Não é permitido o mesmo vínculo para OCT com PAGTO/ADCC ou DEPÓSITO, favor verificar!", vbInformation, App.Title
        Exit Sub
    End If
    
'*********  DESATIVADO TEMPORARIAMENTE A PEDIDO DO PESSOAL DA USB (ASS: Fernando) **********
'    'Verifica se existe Diferença de Valores para Pagamento com CHQ. Terceiro
'    If (iDoctoPagto > 0 Or iDoctoDeposito > 0) And iChqOutros > 0 Then
'        'Verifica se existe diferença de Valores
'        If cTotCreditos <> cTotDebitos Then
'            If (cTotCreditos - cTotDebitos) <> 0 Then
'                MsgBox "Não é permitido víncular Cheque de Terceiro com diferença de valores, favor verificar !", vbInformation, App.Title
'                Exit Sub
'            End If
'        End If
'    End If
    
    '---------------------------------------------------------------------------
    '       Vincula documento à documento na base e atualiza  aDoc(x)
    '---------------------------------------------------------------------------
    For X = Inicio To Fim
        If LstDocto.Selected(X - 1) Then
            Set qryGeraVinculo = Geral.Banco.CreateQuery("", "{? = call GeraVinculoDocumento (?,?,?)}")
            With qryGeraVinculo
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                .rdoParameters(2) = aDoc(X).IdDocto             'IdDocto do documento
                .rdoParameters(3) = lIdDoctoVincular            'IdDocto como vínculo
                .Execute
            End With

            If qryGeraVinculo(0).Value = 1 Then
                MsgBox "Ocorreu um erro ao vincular documento.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            aDoc(X).Vinculo = lIdDoctoVincular
        
            'Documento vinculado manualmente
            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 270)
        
            '-----------------------------------------------------------------------------------
            '       Altera TipoDocto conforme documentos vinculados (Efeito para o Robô)
            '-----------------------------------------------------------------------------------
            If aDoc(X).Status = "1" And (aDoc(X).EstornoDocto = False) Then
                'Verifica se (1) Depósito para (1/n)Cheques
                If iDoctoOCT > 0 Then
                    If aDoc(X).TipoDocto = 5 Or aDoc(X).TipoDocto = 6 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 7) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento transformado em Cheque Depósito
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 274)
                    End If
                ElseIf iDoctoPagto = 0 And iDoctoCheque > 0 And iDoctoDeposito = 1 And iDoctoADCC = 0 Then
                    If aDoc(X).TipoDocto = 5 Or aDoc(X).TipoDocto = 6 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 7) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento transformado em Cheque Depósito
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 274)
                    End If
                
                'Verifica se (n) Depósitos para (1/n)Cheques e se cheque UBB
                ElseIf iDoctoCheque > 0 And iDoctoDeposito >= 1 And iDoctoADCC = 0 Then
                    If aDoc(X).TipoDocto = 6 And InStr("409*230", Left(aDoc(X).Leitura, 3)) <> 0 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 5) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento cheque Compensação para Cheque Sacado
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 275)
                    ElseIf aDoc(X).TipoDocto = 7 And InStr("409*230", Left(aDoc(X).Leitura, 3)) <> 0 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 5) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento cheque Compensação para Cheque Sacado
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 276)
                    ElseIf aDoc(X).TipoDocto = 7 And InStr("409*230", Left(aDoc(X).Leitura, 3)) = 0 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 6) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento cheque Compensação para Cheque Sacado
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 277)
                    End If
                Else
                    If aDoc(X).TipoDocto = 7 And InStr("409*230", Left(aDoc(X).Leitura, 3)) <> 0 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 5) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento cheque Compensação para Cheque Sacado
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 276)
                    ElseIf aDoc(X).TipoDocto = 7 And InStr("409*230", Left(aDoc(X).Leitura, 3)) = 0 Then
                        If Not AlteraTipoDocto(aDoc(X).IdDocto, 6) Then
                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                            Exit For
                        End If
                        'Documento cheque Compensação para Cheque Sacado
                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X).IdDocto, 277)
                    End If
                End If
            End If

        End If
    Next X
    
    Call SetAlteraDocto(True)

    Call PreencheListDocto(LstDocto.ListIndex)

    LstDocto.SetFocus
    Exit Sub

err_cmdVincular_Click:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Ocorreu um erro ao vincular documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select


End Sub

Public Sub cmdZoomMais_Click()

  On Error GoTo ERRO_ZOOMMAIS

  If teclou Then Exit Sub
  
  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  If Lead1.PaintZoomFactor <= 400 Then
      Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
      Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
  End If
  DoEvents
  teclou = False
  Exit Sub

ERRO_ZOOMMAIS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Public Sub cmdZoomMenos_Click()

  On Error GoTo ERRO_ZOOMMENOS

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  If Lead1.PaintZoomFactor >= 20 Then
      Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
      Lead1.PaintZoomFactor = Lead1.PaintZoomFactor - 10
  End If
  DoEvents
  teclou = False
  Exit Sub

ERRO_ZOOMMENOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Private Sub MostraImagem()

  On Error GoTo ERRO_MOSTRAIMAGEM

  Dim Ret As Long

  hCtl = Lead1.hwnd

   'Habilita Objetos de Manipulação de Imagens
    If Not (aDoc(LstDocto.ListIndex + 1).TipoDocto = 32 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 33 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 34 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 38 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 42 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 43 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 44 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 45) Then
        'Coloca imagem na tela
        With Lead1
            If Not .Visible Then .Visible = True
            lblAjuste.Visible = False
            
          .Tag = "F"
          .AutoRepaint = False
          If Geral.VIPSDLL = eDllProservi Then
            .Load Geral.DiretorioImagens & aDoc(LstDocto.ListIndex + 1).Frente, 0, 0, 1
          Else
            .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(LstDocto.ListIndex + 1).Frente, 0, 0, 1
          End If
          
          'Se imagem for da ls500, deixar mais escura
          If aDoc(LstDocto.ListIndex + 1).Ordem <> "2" Then
            .Intensity 220
          Else
            .Intensity 140
          End If
          'Se imagem for do canon, diminui em 50% o tamanho
          If aDoc(LstDocto.ListIndex + 1).Ordem <> "1" Then
            .PaintZoomFactor = 100
          Else
            .PaintZoomFactor = 50
          End If
          .AutoRepaint = True
        End With
    
        FrmImagem.Visible = True
    Else
        If Not lblAjuste.Visible Then lblAjuste.Visible = True
        FrmImagem.Visible = False
        ApresentaTelaAjuste (aDoc(LstDocto.ListIndex + 1).TipoDocto)
    End If

    'Posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)

     'Habilita Objetos de Manipulação de Imagens
    If Not (aDoc(LstDocto.ListIndex + 1).TipoDocto = 32 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 33 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 34 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 38 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 42 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 43 Or _
            aDoc(LstDocto.ListIndex + 1).TipoDocto = 44 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 45) Then
    
        CmdOcorrencia.Enabled = True
        'Habilita Objetos de Manipulação de Imagens
        Call HDObjetosImagem(True)
    Else
        Call HDObjetosImagem(False)
    End If

    DoEvents

Exit Sub

ERRO_MOSTRAIMAGEM:
  Screen.MousePointer = vbDefault
  MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
  Call HDObjetosImagem(False)
  
End Sub

Private Sub Form_Activate()

    On Error GoTo ERRO_ACTIVATE

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(9)
    
    'Desabilita tela se Saldo sobre cheque
    fraSaldo.Visible = False

    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    'Preencher List com as Capas de Documentos para CSP
    If PrimeiraVez Then
        PrimeiraVez = False

        AlterouDocto = False
        
        If Not PreencheListCapas Then
            MsgBox "Não Existem Envelopes / Malotes para C.S.P.", vbInformation, App.Title

            Call HabilitaTimerPesquisa

            Exit Sub
        End If

        sTempo = 0

        'Habilitar o Timer de Atualização
        TmrAtualiza.Enabled = True
    End If

    Exit Sub

ERRO_ACTIVATE:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Ativar Tela.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
Dim Ret As Long

hCtl = frmCSP.Lead1.hwnd

    Select Case KeyCode
        Case vbKeyAdd
            Call cmdZoomMais_Click
        Case vbKeySubtract
            Call cmdZoomMenos_Click
        Case vbKeyF10
            Call cmdInverteCor_Click
            KeyCode = 0
        Case vbKeyDivide
            Call cmdRotacao_Click
        Case vbKeyF11
            Call cmdFrenteVerso_Click
        Case vbKeyDown
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEDOWN, 0)
        Case vbKeyUp
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEUP, 0)
        Case vbKeyLeft
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEUP, 0)
        Case vbKeyRight
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEDOWN, 0)
    End Select

End Sub

Private Sub Form_Load()

    Dim Control As Control
    
    PrimeiraVez = True
    sMasc = "###,###,###,##0.00"
    
    cmdRetiraOcorrencia.Enabled = False
    CmdEnviaCompensacao.Enabled = False
    cmdRemoverVinculo.Enabled = False
    cmdSaldo.Enabled = False
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    'Call GravaLog(0, 0, 260)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    If AlterouDocto Then
'        MsgBox "Foi alterado documento(s) nesta capa, favor verificar e encerrar a capa.", vbInformation, App.Title
'        Cancel = True
'        Exit Sub
'    End If

    'Verificar se foi selecionado uma Capa Anteriormente
    If lstCapa.ListIndex + 1 > 0 Then
        If aCapa(lstCapa.ListIndex + 1).Status <> "V" Then
            Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "N")
            aCapa(lstCapa.ListIndex + 1).Status = "N"
        End If
    End If

    IdSelecionado = 0
    
    'Desabilitar os timers
    TmrAtualiza.Enabled = False
    TmrPesquisa.Enabled = False
    
    'Finalizar Conexões
    Set qryGetCapa = Nothing
    Set qryGetDocumentos = Nothing
    Set qryAtualizaStatusCapa = Nothing
    Set qryGetOcorr = Nothing
    Set qryAtualizaOcorrencia = Nothing

    Set qryRemoveAjusteCapa = Nothing
    Set qryEnviarCompensacao = Nothing
    Set qryGeraVinculo = Nothing
    Set qryControleCapa = Nothing
    Set qryInsereAjuste = Nothing
    Set qryAlteraTipoDocto = Nothing
    Set qryGetUltimaOrdemCaptura = Nothing
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    'Call GravaLog(0, 0, 261)

  
End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 1 Then
    Lead1.AutoRubberBand = True
    Lead1.MousePointer = vbCrosshair
  Else
    Call MostraImagem
  End If
End Sub
Private Sub Lead1_RubberBand()

  On Error GoTo ERRO_RUBBERBAND

  Dim zoomleft As Integer
  Dim zoomtop As Integer
  Dim zoomwidth As Integer
  Dim zoomheight As Integer

  Lead1.MousePointer = 0
  'Zoom in on the selection.
  zoomleft = Lead1.RubberBandLeft
  zoomtop = Lead1.RubberBandTop
  zoomwidth = Lead1.RubberBandWidth
  zoomheight = Lead1.RubberBandHeight
  If (zoomwidth = 0) Or (zoomheight = 0) Then
      Exit Sub
  End If

  'Zoom in on the rectangle defined by the rubberband
  Lead1.ZoomToRect zoomleft, zoomtop, zoomwidth, zoomheight
  Lead1.ForceRepaint

  Exit Sub

ERRO_RUBBERBAND:
  MsgBox "Não é possível redimensionar a Imagem.", vbInformation, App.Title
End Sub

Private Sub lstCapa_Click()

Dim rsDocumentos As rdoResultset
Dim sSql As String
Dim X As Integer
Dim sLinha As String
Dim Ret As Integer
Dim Status As String

On Error GoTo ERRO_CAPACLICK

    If fraSaldo.Visible Then
        Call TelaConsultaSaldo(False)
    End If
    
    If Screen.MousePointer = vbDefault And lstCapa.ListIndex <> -1 Then
        
        Screen.MousePointer = vbHourglass
    
        AlterouDocto = False
        cmdVincular.Enabled = False
'''        cmdEstorno.Enabled = False
        
        sTempo = 0
    
        If IdSelecionado <> 0 And (IdSelecionado <> aCapa(lstCapa.ListIndex + 1).IdCapa) Then
            'A Capa anterior não sofreu alteração , Voltar o Status para 'N'
            Call AtualizaStatusCapa(IdSelecionado, "N")
        End If
    
        'Verificar se a Capa está Duplicada
        If CapaDuplicada Then
            bCapaDuplicada = True
            LblCapaDup.Visible = True
        Else
            bCapaDuplicada = False
            LblCapaDup.Visible = False
        End If
    
        'Verificar se a capa mudou
        If IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa Then
            Call PreencheListDocto(0)
        Else
            'Obtem capa disponível para CSP
            Ret = CapaSelecionadaDisponivel
            If Ret = 0 Then
                
'                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "Q")
                aCapa(lstCapa.ListIndex + 1).Status = "Q"
            
                'Grava log - Selecionar Capa'
                'GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 269
        
                Call PreencheListDocto(0)
            Else
                Call HDObjetosImagem(False)
                IdSelecionado = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
        'Verificar se a Capa é 'ENVELOPE' OU 'MALOTE'
        If UCase(Trim(aCapa(lstCapa.ListIndex + 1).IdEnv_Mal)) = "E" Then
            'Envelope
            LblEnv_Mal.Caption = "Envelope"
    
            Call HDMalote(False)
        Else
            'Malote
            LblEnv_Mal.Caption = "Malote"
            
            Call HDMalote(True)
            lblNumMalote.Caption = aCapa(lstCapa.ListIndex + 1).NumMalote
        End If
    
        'Informar o Lote na tela
        lblLote.Caption = Format(Trim(aCapa(lstCapa.ListIndex + 1).IdLote), "0000-00000")
    
        'Limpar Objetos
        lblOcorrencia.Caption = ""
    
        Screen.MousePointer = vbDefault
        
        Call SetAlteraDocto(AlterouDocto)
    End If
    
    IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa
    
Exit Sub

ERRO_CAPACLICK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Capa do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Function ExcluiAjuste() As Boolean

  On Error GoTo Erro_ExcluiAjuste

  ExcluiAjuste = False

  Set qryRemoveAjusteCapa = Geral.Banco.CreateQuery("", "{? = call RemoveAjusteCapaCsp (?,?)}")
  With qryRemoveAjusteCapa
    .rdoParameters(0).Direction = rdParamReturnValue
    .rdoParameters(1) = Geral.DataProcessamento                   'Data Proc.
    .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa       'IdCapa
    .Execute
  End With

  If qryRemoveAjusteCapa(0).Value = 1 Then
    'Ocorreu um erro
    MsgBox "Ocorreu um erro ao excluir Ajustes.", vbInformation + vbOKOnly, App.Title
    Exit Function
  End If

  ExcluiAjuste = True

  Exit Function

Erro_ExcluiAjuste:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Ajustes na Capa.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select

End Function
Private Sub lstDocto_Click()

   On Error GoTo ERRO_DOCTOCLICK

   Dim RsOcorr As rdoResultset
   Dim sSql As String
   Dim X As Integer
   Dim sOcorrencia As String

   'Exibir a Figura do Documento Selecionado
   Call MostraImagem

   lblOcorrencia.Caption = ""

   'Verifica se o Documento possui Ocorrência
   If aDoc(LstDocto.ListIndex + 1).Status = "D" Or _
        aDoc(LstDocto.ListIndex + 1).Status = "C" Or _
        aDoc(LstDocto.ListIndex + 1).Status = "F" Then
        
        'Verificar se a ocorrencia começa com 999
        If Left(aDoc(LstDocto.ListIndex + 1).Ocorrencia, 3) = "999" Then
        
            If aDoc(LstDocto.ListIndex + 1).RetornoTransacao > 0 Then
                Call ObtemRetornoTransacao(aDoc(LstDocto.ListIndex + 1).RetornoTransacao, sOcorrencia)
            Else
                sOcorrencia = "Erro operacional."
            End If
            lblOcorrencia.Caption = sOcorrencia
        
        Else
            'Verificar se o código da ocorrencia possui 3 ou 5 caracteres
            If Val(aDoc(LstDocto.ListIndex + 1).Ocorrencia) > 999 Then
               '5 Posicoes
               sSql = Left(Trim(aDoc(LstDocto.ListIndex + 1).Ocorrencia), 3)
            Else
               '3 Posicoes
               If Right(Trim(aDoc(LstDocto.ListIndex + 1).Ocorrencia), 2) = "00" Then
                  'Ocorrencia atualizada pelo robo
                  sSql = Val(Trim(aDoc(LstDocto.ListIndex + 1).Ocorrencia)) / 100
               Else
                  'Ocorrencia gerada pelo sistema
                  sSql = Val(Trim(aDoc(LstDocto.ListIndex + 1).Ocorrencia))
               End If
            End If
            
            Set qryGetOcorr = Geral.Banco.CreateQuery("", "{call GetOcorrencia (" & sSql & ")}")
            
            Set RsOcorr = qryGetOcorr.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            
            lblOcorrencia.Caption = ""
            If Not RsOcorr.EOF Then
               lblOcorrencia.Caption = "Ocorrência : " & RsOcorr!Descricao
            End If
        End If
       
   End If

    'Se documento com ocorrência (gerada pelo robô), habilita opção de retirar ocorrência
    If (aDoc(LstDocto.ListIndex + 1).Status = "F" Or aDoc(LstDocto.ListIndex + 1).Status = "C") And _
        aDoc(LstDocto.ListIndex + 1).TipoDocto <> 1 And _
        aDoc(LstDocto.ListIndex + 1).Ocorrencia <> 0 Then
        cmdRetiraOcorrencia.Enabled = True
    Else
        cmdRetiraOcorrencia.Enabled = False
    End If
    
    If aDoc(LstDocto.ListIndex + 1).TipoDocto = 1 Or _
        aDoc(LstDocto.ListIndex + 1).Status = "D" Or _
        aDoc(LstDocto.ListIndex + 1).Status = "T" Or _
        aDoc(LstDocto.ListIndex + 1).EstornoDocto Or _
        aDoc(LstDocto.ListIndex + 1).Duplicidade Then
        
        If LstDocto.Selected(LstDocto.ListIndex) Then
            CmdOcorrencia.Enabled = False
        Else
            CmdOcorrencia.Enabled = True
        End If
    Else
        CmdOcorrencia.Enabled = True
    End If
    
    'Permitir enviar cheque UBB Sacado para compensação
    If aDoc(LstDocto.ListIndex + 1).Status = "1" And (aDoc(LstDocto.ListIndex + 1).TipoDocto = 5 Or aDoc(LstDocto.ListIndex + 1).TipoDocto = 7) And aDoc(LstDocto.ListIndex + 1).EstornoDocto = False Then
'    If InStr("F,C,1", aDoc(LstDocto.ListIndex + 1).Status) <> 0 And aDoc(LstDocto.ListIndex + 1).TipoDocto = 5 And aDoc(LstDocto.ListIndex + 1).EstornoDocto = False Then
        CmdEnviaCompensacao.Enabled = True
    Else
        CmdEnviaCompensacao.Enabled = False
    End If

    'Permitir retirar vínculo
    If aDoc(LstDocto.ListIndex + 1).Status = "1" And aDoc(LstDocto.ListIndex + 1).EstornoDocto = False And aDoc(LstDocto.ListIndex + 1).TipoDocto <> "1" Then
        cmdRemoverVinculo.Enabled = True
    Else
        If LstDocto.Selected(LstDocto.ListIndex) Then
            cmdRemoverVinculo.Enabled = False
        Else
            cmdRemoverVinculo.Enabled = True
        End If
    End If

    'Permitir vínculo para documento não transmitido
'    If InStr("T*F*D*C", aDoc(LstDocto.ListIndex + 1).Status) <> 0 Then
    If InStr("*F*D*C", aDoc(LstDocto.ListIndex + 1).Status) <> 0 Then
        If LstDocto.Selected(LstDocto.ListIndex) Then
            cmdVincular.Enabled = False
        Else
            cmdVincular.Enabled = True
        End If
    Else
        cmdVincular.Enabled = True
    End If

    'Permitir consultar Saldo para cheque UBB
    If LstDocto.SelCount = 1 And _
        InStr("5,6", aDoc(LstDocto.ListIndex + 1).TipoDocto) <> 0 And InStr("409,230", Left(aDoc(LstDocto.ListIndex + 1).Leitura, 3)) <> 0 Then
        cmdSaldo.Enabled = True
    Else
        cmdSaldo.Enabled = False
    End If

    Exit Sub

ERRO_DOCTOCLICK:
    
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Selecionar Documento.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
End Sub


Private Sub lstDocto_DblClick()
    
    If cmdSaldo.Enabled Then
        Call cmdSaldo_Click
    End If
    
End Sub

Private Sub tmrAtualiza_Timer()

    TmrAtualiza.Enabled = False
    
    If lstCapa.ListIndex <> -1 Then
        If aCapa(lstCapa.ListIndex + 1).IdCapa <> 0 Then
            sTempo = sTempo + Int(TmrAtualiza.Interval / 1000)
            If sTempo + Int(TmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
                'Atualizar o Status da Capa
                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "Q")
                sTempo = 0
            End If
        End If
    End If
    
    TmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()

  TmrPesquisa.Enabled = False

  sTempo = sTempo + Int(TmrPesquisa.Interval / 1000)

  If sTempo + Int(TmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    'Pesquisar por Documentos para C.S.P.
    sTempo = 0

    If PreencheListCapas Then
        'Call GravaLog(0, 0, 268)        'Grava log (CSP - Fim Aguarda documento)
        Exit Sub
    End If

    TmrPesquisa.Enabled = True
  End If

  'Atualizar a Barra de Progresso
  If Progress.Value + 4 > 100 Then
    Progress.Value = 0
  Else
    Progress.Value = Progress.Value + 4
  End If

  DoEvents
  TmrPesquisa.Enabled = True
End Sub

Private Sub TxtNumEnvMal_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call cmdProcurar_Click
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub

Private Sub AtualizaCtrlCapa(IdCapa As Long, Comentarios As String, IdModulo As Long)

On Error GoTo Err_AtualizaCtrlCapa

    Screen.MousePointer = vbHourglass

    'Atualizar o registro controle de capa
    Set qryControleCapa = Geral.Banco.CreateQuery("", "{? = call InsereControleCapa (?,?,?,?)}")
    With qryControleCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = IdCapa                      'IdCapa
        .rdoParameters(3) = Comentarios                 'IdDocto
        .rdoParameters(4) = IdModulo                    'Modulo
        .Execute
    End With

    If qryControleCapa(0).Value = 1 Then
        MsgBox "Ocorreu um erro ao atualizar o controle de capa.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbDefault

    Exit Sub
    
Err_AtualizaCtrlCapa:
  
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao atualizar Controle de Capa.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select

End Sub

Private Function SomaDebitosCreditos(lVinculo As Long, cTotCreditos As Currency, cTotDebitos As Currency, ByRef bSomenteDeposito As Boolean)
'   Somentedeposito -   Parâmetro de retorno onde irá informar se para este vínculo
'                       existe somente DEPÓSITO/OCT com cheques sem nenhuma conta ou algo diferente disso
Dim X As Integer

bSomenteDeposito = True

    For X = 0 To LstDocto.ListCount - 1
        If InStr("T*1", aDoc(X + 1).Status) <> 0 Then
        
            If aDoc(X + 1).Vinculo = lVinculo And (aDoc(X + 1).EstornoDocto = False) Then
                Select Case aDoc(X + 1).TipoDocto
                    Case 4, 5, 6, 7, 41
                        cTotCreditos = cTotCreditos + aDoc(X + 1).Valor
                    Case 32, 34, 42, 44 'Soma-se como Débito para Contra partida
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 2, 3, 37      ' Depositos e OCT
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 8 To 31, 35, 36, 40   'Pagtos
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 33, 38, 43, 45 'Soma-se como Crédito para Contra partida
                        cTotCreditos = cTotCreditos + aDoc(X + 1).Valor
                End Select
                'Verifica se vínculo somente com depósito
                If Not (aDoc(X + 1).TipoDocto = 39 Or aDoc(X + 1).TipoDocto = 37 Or _
                    aDoc(X + 1).TipoDocto = 2 Or aDoc(X + 1).TipoDocto = 3 Or _
                    aDoc(X + 1).TipoDocto = 5 Or aDoc(X + 1).TipoDocto = 6 Or aDoc(X + 1).TipoDocto = 7) Then
                    bSomenteDeposito = False
                End If
            End If
        End If
    Next

End Function

Private Function DigitaAgenciaConta(ByVal lVinculo, ByRef Agencia As Integer, ByRef Conta As Long, ByVal cValorDiferenca As Currency) As Boolean

    On Error GoTo Erro_DigitaAgenciaConta

    DigitaAgenciaConta = False
    
    AgenciaContaAjuste.m_Vinculo = lVinculo
    AgenciaContaAjuste.m_Diferenca = cValorDiferenca
    Call AgenciaContaAjuste.ShowModal(Agencia, Conta)
    AgenciaContaAjuste.m_Diferenca = 0

    If (Agencia <> 0 And Conta <> 0) Then
        DigitaAgenciaConta = True
    End If

    Exit Function

Erro_DigitaAgenciaConta:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do vínculo do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function

Private Function GeraAjuste() As Boolean

Dim rsOrdemCaptura As rdoResultset
Dim strAutenticacaoDigital As String

On Error GoTo ErroAjuste

    GeraAjuste = False

    'Verificar qual o ultimo numero de ordem de captura e incrementar 1
    Set qryGetUltimaOrdemCaptura = Geral.Banco.CreateQuery("", "{Call GetUltimaOrdemCaptura (?,?)}")
    qryGetUltimaOrdemCaptura.rdoParameters(0).Value = lstCapa.ItemData(lstCapa.ListIndex)
    qryGetUltimaOrdemCaptura.rdoParameters(1).Value = Geral.DataProcessamento
    
    Set rsOrdemCaptura = qryGetUltimaOrdemCaptura.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    Set qryInsereAjuste = Geral.Banco.CreateQuery("", "{ ? = Call InsereAjuste (?,?,?,?,?,?,?,?,?)}")

    'Gera Autenticação Digital
    strAutenticacaoDigital = G_EncriptaBO(tpAjuste.TipoDocto, CStr(tpAjuste.Conta))
    If strAutenticacaoDigital = "" Then GoTo ErroAjuste

    With qryInsereAjuste
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lstCapa.ItemData(lstCapa.ListIndex)
        .rdoParameters(3) = tpAjuste.TipoDocto
        .rdoParameters(4) = tpAjuste.Agencia
        .rdoParameters(5) = tpAjuste.Conta
        .rdoParameters(6) = tpAjuste.Valor
        .rdoParameters(7) = tpAjuste.Vinculo
        .rdoParameters(8) = Val(rsOrdemCaptura!MaiorOrdem) + 1
        .rdoParameters(9) = strAutenticacaoDigital
        .Execute

        If .rdoParameters(0) <> 0 Then
            GoTo ErroAjuste
        End If
    End With

    If Not (rsOrdemCaptura Is Nothing) Then rsOrdemCaptura.Close
    
    GeraAjuste = True
    Exit Function
    
ErroAjuste:
    
    Beep
    If Not (rsOrdemCaptura Is Nothing) Then rsOrdemCaptura.Close
    
    Select Case TratamentoErro("Erro na inserção de ajuste de credito/debito.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Unload Me
    Exit Function

End Function
Private Sub SetAlteraDocto(ByVal bAlterou As Boolean)

    AlterouDocto = bAlterou
'    CmdAtualizar.Enabled = Not bAlterou
'    CmdLocalizar.Enabled = Not bAlterou
    
'    lstCapa.Enabled = Not bAlterou
    
End Sub

Private Function CapaComDiferenca(rsDocumentos As rdoResultset) As Boolean

Dim nReg As Long
Dim Vinculo As Long
Dim cTotalCreditos As Currency, cTotalDebitos As Currency

On Error GoTo err_CapaComDiferenca

    CapaComDiferenca = False
    
    Vinculo = 0
    rsDocumentos.MoveFirst
        
    While Not rsDocumentos.EOF
        'Localiza documentos diferentes de transmitidos

        If rsDocumentos!Status <> "T" Then
            If rsDocumentos!TipoDocto <> 1 And rsDocumentos!Vinculo <> 0 Then
    
                If Vinculo <> rsDocumentos!Vinculo Then
                    Vinculo = rsDocumentos!Vinculo
                    cTotalCreditos = 0
                    cTotalDebitos = 0
                    nReg = rsDocumentos.AbsolutePosition
                    
                    rsDocumentos.MoveFirst
                    
                    'Soma Débitos e Créditos por vínculo
                    While Not rsDocumentos.EOF
                        If (rsDocumentos!Status = "1" Or rsDocumentos!Status = "T") _
                            And rsDocumentos!TipoDocto <> 1 And _
                            IsNull(rsDocumentos!IdDoctoEstorno) Then
                        
                            If rsDocumentos!Vinculo = Vinculo Then
                                Select Case rsDocumentos!TipoDocto
                                    Case 4, 5, 6, 7, 41
                                        cTotalCreditos = cTotalCreditos + rsDocumentos!Valor
                                    Case 32, 34, 42, 44 'Contra partida
                                        cTotalDebitos = cTotalDebitos + rsDocumentos!Valor
                                    Case 2, 3, 37      ' Depositos e OCT
                                        cTotalDebitos = cTotalDebitos + rsDocumentos!Valor
                                    Case 8 To 31, 35, 36, 40 'Pagtos
                                        cTotalDebitos = cTotalDebitos + rsDocumentos!Valor
                                    Case 33, 38, 43, 45 'Contra Partida
                                        cTotalCreditos = cTotalCreditos + rsDocumentos!Valor
                                End Select
                            End If
                        End If
                        rsDocumentos.MoveNext
                    Wend
                    
                    If cTotalCreditos <> cTotalDebitos Then
                        CapaComDiferenca = True
                        Exit Function
                    End If
                    
                    rsDocumentos.AbsolutePosition = nReg
                End If
            End If

        End If
        
        rsDocumentos.MoveNext
    Wend
    
    Exit Function
    
err_CapaComDiferenca:
    Screen.MousePointer = vbDefault
    
    Select Case TratamentoErro("Erro ao verificar diferença de valores nos documentos.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    
End Function
Private Function ExisteDoctoSemVinculo() As Boolean

Dim X As Integer
    
    ExisteDoctoSemVinculo = False

    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).Status = "1" And aDoc(X + 1).TipoDocto <> 1 Then
        
            If aDoc(X + 1).Vinculo = 0 Then
                ExisteDoctoSemVinculo = True
                Exit Function
            End If
        End If
    Next

End Function
Private Function ExisteDoctoSemEstorno() As Boolean

Dim X As Integer
    
    ExisteDoctoSemEstorno = False

    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).TipoDocto <> 1 Then
            If Not aDoc(X + 1).EstornoDocto Then
                ExisteDoctoSemEstorno = True
                Exit Function
            End If
        End If
    Next

End Function
Private Sub TelaConsultaSaldo(bHabilita As Boolean)

    fraSaldo.Visible = bHabilita
    FraCmd.Enabled = Not bHabilita
    LstDocto.Enabled = Not bHabilita

End Sub
Private Sub ApresentaTelaAjuste(ByVal intTipoDocto As Integer)

    'Centraliza Label de Ajuste
    lblAjuste.Top = FrmImagem.Top + (FrmImagem.Height - lblAjuste.Height) / 2
    lblAjuste.Left = FrmImagem.Left + (FrmImagem.Width - lblAjuste.Width) / 2
    
     'Habilita Objetos de Manipulação de Imagens
    If intTipoDocto = 32 Or intTipoDocto = 34 Or intTipoDocto = 42 Or intTipoDocto = 44 Then
        lblAjuste.Caption = "Ajuste Crédito"
    ElseIf intTipoDocto = 33 Or intTipoDocto = 38 Or intTipoDocto = 43 Or intTipoDocto = 45 Then
        lblAjuste.Caption = "Ajuste Débito"
    End If
    
End Sub
Private Function ExisteDoctoUnicoVinculo() As Long
'----------------------------------------------------------------
'   Verifica em todos documento se existe único documento com o
'   mesmo número de vínculo
'----------------------------------------------------------------
Dim avinculo()  As Long
Dim Vinculo As Long
Dim X As Integer, Y As Integer
Dim bVerificado As Boolean
Dim soma As Integer
Dim iVinculo As Integer

    Vinculo = 0
    ExisteDoctoUnicoVinculo = 0
    ReDim avinculo(0)
    
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).TipoDocto <> 1 And aDoc(X + 1).Vinculo <> 0 Then

            If Vinculo <> aDoc(X + 1).Vinculo Then
                Vinculo = aDoc(X + 1).Vinculo
                
                'Verifica se vinculo já sofreu verificação
                bVerificado = False
                For iVinculo = 1 To UBound(avinculo)
                    If avinculo(iVinculo) = Vinculo Then
                        bVerificado = True
                        Exit For
                    End If
                Next
                
                If Not bVerificado Then
                    ReDim Preserve avinculo(UBound(avinculo) + 1)
                    avinculo(UBound(avinculo)) = Vinculo
                End If
                
                If Not bVerificado Then
                    soma = 0
                    For Y = 0 To LstDocto.ListCount - 1
                        If InStr("T*1", aDoc(Y + 1).Status) <> 0 Then
                            If aDoc(Y + 1).Vinculo = Vinculo And (aDoc(Y + 1).EstornoDocto = False) Then
                                soma = soma + 1
                            End If
                        End If
                    Next
                    If soma = 1 Then
                        ExisteDoctoUnicoVinculo = Vinculo
                        Exit Function
                    End If
                End If
            
            End If
        End If
    Next

End Function
Private Function ExisteContraPartidaPorVinculo() As Boolean
'-----------------------------------------------------------------------------
'   Verifica Vínculo à Vínculo se existem as devidas contas Débito e Crédito
'   e se está dentro da regra de desdobro
'-----------------------------------------------------------------------------
Dim avinculo()  As Long
Dim Vinculo As Long
Dim X As Integer, Y As Integer
Dim bVerificado As Boolean
Dim iVinculo As Integer
Dim cTotalCred As Currency, cTotalDeb As Currency

'Acumuladores para soma de documento do mesmo tipo
Dim iDoctoDeposito As Integer, iDoctoOCT As Integer, iDoctoLanctoInterno As Integer
Dim iDoctoADCC As Integer, iDoctoCheque As Integer, iDoctoPagto As Integer
Dim iChqUBB As Integer, iChqOutros As Integer

    Vinculo = 0
    ExisteContraPartidaPorVinculo = False
    ReDim avinculo(0)
    
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).TipoDocto <> 1 And aDoc(X + 1).Vinculo <> 0 Then

            If Vinculo <> aDoc(X + 1).Vinculo Then
                Vinculo = aDoc(X + 1).Vinculo
                
                'Verifica se vinculo já sofreu verificação
                bVerificado = False
                For iVinculo = 1 To UBound(avinculo)
                    If avinculo(iVinculo) = Vinculo Then
                        bVerificado = True
                        Exit For
                    End If
                Next
                
                If Not bVerificado Then
                    ReDim Preserve avinculo(UBound(avinculo) + 1)
                    avinculo(UBound(avinculo)) = Vinculo
                End If
                
                If Not bVerificado Then
                    iDoctoDeposito = 0: iDoctoOCT = 0:      iDoctoLanctoInterno = 0
                    iDoctoADCC = 0:     iDoctoCheque = 0:   iDoctoPagto = 0
                    iChqUBB = 0:        iChqOutros = 0
                    
                    For Y = 0 To LstDocto.ListCount - 1
                        If InStr("T*1", aDoc(Y + 1).Status) <> 0 Then
                            If aDoc(Y + 1).Vinculo = Vinculo And (aDoc(Y + 1).EstornoDocto = False) Then
                                Select Case aDoc(Y + 1).TipoDocto
                                    Case 2, 3           ' Depositos
                                        iDoctoDeposito = iDoctoDeposito + 1
                                    Case 4              ' ADCC
                                        iDoctoADCC = iDoctoADCC + 1
                                    Case 5, 6, 7        ' Cheques
                                        iDoctoCheque = iDoctoCheque + 1
                                        'Verifica se cheque UBB
                                        If InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) <> 0 Then
                                            iChqUBB = iChqUBB + 1
                                        Else
                                            iChqOutros = iChqOutros + 1
                                        End If
                                    Case 37             ' OCT
                                        iDoctoOCT = iDoctoOCT + 1
                                    Case 41             ' Lancamento Interno
                                        iDoctoLanctoInterno = iDoctoLanctoInterno + 1
                                    Case 39, 32, 33, 34, 38, 42, 43, 44, 45
                                        'Capa de OCT e Ajustes não tratar
                                    Case Else                ' Doctos para Pagamento
                                        iDoctoPagto = iDoctoPagto + 1
                                End Select
                            End If
                        End If
                    Next
                    
                    '---------------------------------------------------------------------------
                    '       Verifica se existe documento débito e crédito para vínculo
                    '---------------------------------------------------------------------------
                    If iDoctoDeposito > 0 And iDoctoLanctoInterno = 0 And iDoctoCheque = 0 And iDoctoADCC = 0 Then
                        MsgBox "Depósito sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoCheque > 0 And iDoctoPagto = 0 And iDoctoOCT = 0 And iDoctoDeposito = 0 Then
                        MsgBox "Cheque(s) sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoOCT > 0 And iDoctoLanctoInterno = 0 And iDoctoCheque = 0 And iDoctoADCC = 0 Then
                        MsgBox "OCT sem cheque para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoADCC > 0 And iDoctoPagto = 0 And iDoctoOCT = 0 And iDoctoDeposito = 0 Then
                        MsgBox "ADCC sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                        
                    If iDoctoLanctoInterno > 0 And iDoctoPagto = 0 And iDoctoDeposito = 0 And iDoctoOCT = 0 Then
                        MsgBox "Lançamento Interno sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoPagto > 0 And iDoctoCheque = 0 And iDoctoLanctoInterno = 0 And iDoctoADCC = 0 Then
                        MsgBox "Pagamento sem documento para efetivar o vínculo, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    '---------------------------------------------------------------------------------------
                    '                   Verifica regra para desdobramento de cheque
                    '---------------------------------------------------------------------------------------
                
                    If iDoctoPagto > 0 And iDoctoCheque > 0 And iDoctoDeposito > 0 And iChqOutros > 0 Then
                        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoCheque > 0 And iDoctoDeposito > 1 And iChqOutros > 0 Then
                        MsgBox "Não é permitido o mesmo vínculo para Cheque Outros Bancos com desdobramento, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoCheque > 0 And iDoctoDeposito > 0 And iDoctoADCC > 0 Then
                        MsgBox "Não é permitido o mesmo vínculo para Cheque com desdobramento , favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                    If iDoctoOCT > 0 And (iDoctoPagto > 0 Or iDoctoDeposito > 0 Or iDoctoADCC > 0) Then
                        MsgBox "Não é permitido o mesmo vínculo para OCT com PAGTO/ADCC ou DEPÓSITO, favor verificar!", vbInformation, App.Title
                        Exit Function
                    End If
                    
                End If
            
'*********  DESATIVADO TEMPORARIAMENTE A PEDIDO DO PESSOAL DA USB (ASS: Fernando) **********
'                'Verifica se existe Diferença de Valores para Pagamento com CHQ. Terceiro
'                If (iDoctoPagto > 0 Or iDoctoDeposito > 0) And iChqOutros > 0 Then
'                    cTotalCred = 0
'                    cTotalDeb = 0
'                    Call SomaDebitosCreditos(Vinculo, cTotalCred, cTotalDeb)
'                    'Verifica se existe diferença de Valores
'                    If cTotalCred <> cTotalDeb Then
'                        If (cTotalCred - cTotalDeb) <> 0 Then
'                            MsgBox "Não é permitido víncular Cheque de Terceiros com diferença de valores, favor verificar !" & vbCrLf & vbCrLf & "Vínculo Nr. " & CStr(Vinculo), vbInformation, App.Title
'                            Exit Function
'                        End If
'                    End If
'                End If
            
            End If
        End If
    Next
    
    ExisteContraPartidaPorVinculo = True
    
End Function
Private Function AlteraTipoDocto(ByVal lngIdDocto As Long, ByVal intTipoDocto As Integer) As Boolean

On Error GoTo Err_AlteraTipoDocto

    AlteraTipoDocto = False

    'Verificar qual o ultimo numero de ordem de captura e incrementar 1
    Set qryAlteraTipoDocto = Geral.Banco.CreateQuery("", " { ? = Call AlteraTipoDocto (?,?,?)}")
    With qryAlteraTipoDocto
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = lngIdDocto
        .rdoParameters(3).Value = intTipoDocto
        .Execute

        If .rdoParameters(0) <> 0 Then
            GoTo Err_AlteraTipoDocto
        End If
    End With

    AlteraTipoDocto = True
    Exit Function
    
Err_AlteraTipoDocto:
    
End Function

Private Function AcertaTipoDocto() As Boolean
'-----------------------------------------------------------------------------
'   Verifica Vínculo à Vínculo se existem Depósitos com cheque onde o tipodocto
'   esteja fora da regra
'
'   Para (1) DEP mudar para CHQ Depósito Tipo(7)
'   Para (n) DEP e Cheque UBB, mudar para CHQ UBB Sacado Tipo(5)
'   Para (n) DEP com CHQ UBB e Mais PGTO, mudar para CHQ UBB Sacado Tipo(5)
'-----------------------------------------------------------------------------
Dim avinculo()  As Long
Dim Vinculo As Long
Dim X As Integer, Y As Integer
Dim bVerificado As Boolean
Dim iVinculo As Integer

'Acumuladores para soma de documento do mesmo tipo
Dim iDoctoDeposito As Integer, iDoctoOCT As Integer, iDoctoLanctoInterno As Integer
Dim iDoctoADCC As Integer, iDoctoCheque As Integer, iDoctoPagto As Integer

    Vinculo = 0
    AcertaTipoDocto = False
    ReDim avinculo(0)
    
    For X = 0 To LstDocto.ListCount - 1
        If aDoc(X + 1).TipoDocto <> 1 And aDoc(X + 1).Vinculo <> 0 Then

            If Vinculo <> aDoc(X + 1).Vinculo Then
                Vinculo = aDoc(X + 1).Vinculo
                
                'Verifica se vinculo já sofreu verificação
                bVerificado = False
                For iVinculo = 1 To UBound(avinculo)
                    If avinculo(iVinculo) = Vinculo Then
                        bVerificado = True
                        Exit For
                    End If
                Next
                
                If Not bVerificado Then
                    ReDim Preserve avinculo(UBound(avinculo) + 1)
                    avinculo(UBound(avinculo)) = Vinculo
                End If
                
                'Soma-se todos doctos por vínculo caso ainda não tenha sido feito
                If Not bVerificado Then
                    iDoctoDeposito = 0: iDoctoOCT = 0:      iDoctoLanctoInterno = 0
                    iDoctoADCC = 0:     iDoctoCheque = 0:   iDoctoPagto = 0
                    
                    For Y = 0 To LstDocto.ListCount - 1
                        If InStr("T*1", aDoc(Y + 1).Status) <> 0 Then
                            If aDoc(Y + 1).Vinculo = Vinculo And (aDoc(Y + 1).EstornoDocto = False) Then
                                Select Case aDoc(Y + 1).TipoDocto
                                    Case 2, 3           ' Depositos
                                        iDoctoDeposito = iDoctoDeposito + 1
                                    Case 4              ' ADCC
                                        iDoctoADCC = iDoctoADCC + 1
                                    Case 5, 6, 7        ' Cheques
                                        iDoctoCheque = iDoctoCheque + 1
                                    Case 37             ' OCT
                                        iDoctoOCT = iDoctoOCT + 1
                                    Case 41             ' Lancamento Interno
                                        iDoctoLanctoInterno = iDoctoLanctoInterno + 1
                                    Case 39                 'Capa de OCT não tratar
                                    Case Else                ' Doctos para Pagamento
                                        iDoctoPagto = iDoctoPagto + 1
                                End Select
                            End If
                        End If
                    Next
                    
                    '-----------------------------------------------------------------------------------
                    '     Altera TipoDocto conforme documentos vinculados (Efeito para o Robô)
                    '-----------------------------------------------------------------------------------
                    For Y = 0 To LstDocto.ListCount - 1
                        If aDoc(Y + 1).Status = "1" Then
                            If aDoc(Y + 1).Vinculo = Vinculo And (aDoc(Y + 1).EstornoDocto = False) Then
                    
                                If iDoctoOCT > 0 Then
                                    If aDoc(Y + 1).TipoDocto = 5 Or aDoc(Y + 1).TipoDocto = 6 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 7) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        aDoc(Y + 1).TipoDocto = 7
                                        'Documento transformado em Cheque Depósito
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 274)
                                    End If
                                    
                                'Verifica se (1) Depósito para (1/n)Cheques
                                ElseIf iDoctoPagto = 0 And iDoctoCheque > 0 And iDoctoDeposito = 1 And iDoctoADCC = 0 Then
                                    If aDoc(Y + 1).TipoDocto = 5 Or aDoc(Y + 1).TipoDocto = 6 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 7) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        aDoc(Y + 1).TipoDocto = 7
                                        'Documento transformado em Cheque Depósito
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 274)
                                    End If
                                'Verifica se (n) Depósitos para (1/n)Cheques e se cheque UBB
                                ElseIf iDoctoCheque > 0 And iDoctoDeposito >= 1 And iDoctoADCC = 0 Then
                                    If aDoc(Y + 1).TipoDocto = 6 And InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) <> 0 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 5) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        'Documento cheque Compensação para Cheque Sacado
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 275)
                                    ElseIf aDoc(Y + 1).TipoDocto = 7 And InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) <> 0 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 5) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        'Documento cheque Compensação para Cheque Sacado
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 276)
                                    ElseIf aDoc(Y + 1).TipoDocto = 7 And InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) = 0 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 6) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        'Documento cheque Compensação para Cheque Sacado
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 277)
                                    End If
                                Else
                                    If aDoc(Y + 1).TipoDocto = 7 And InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) <> 0 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 5) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        'Documento cheque Compensação para Cheque Sacado
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 276)
                                    ElseIf aDoc(Y + 1).TipoDocto = 7 And InStr("409*230", Left(aDoc(Y + 1).Leitura, 3)) = 0 Then
                                        If Not AlteraTipoDocto(aDoc(Y + 1).IdDocto, 6) Then
                                            MsgBox "Erro na atualização do vínculo, favor verificar! ", vbCritical + vbOKOnly, App.Title
                                            Exit For
                                        End If
                                        'Documento cheque Compensação para Cheque Sacado
                                        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(Y + 1).IdDocto, 277)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    AcertaTipoDocto = True

End Function
Private Function ExisteDoctoPendenteNaoTransmitido() As Boolean

Dim X As Integer
    
    ExisteDoctoPendenteNaoTransmitido = True

    For X = 1 To LstDocto.ListCount
        If aDoc(X).TipoDocto <> 1 Then
            If Not (aDoc(X).TipoDocto = 32 Or aDoc(X).TipoDocto = 33 Or _
                    aDoc(X).TipoDocto = 34 Or aDoc(X).TipoDocto = 38 Or _
                    aDoc(X).TipoDocto = 42 Or aDoc(X).TipoDocto = 43 Or _
                    aDoc(X).TipoDocto = 44 Or aDoc(X).TipoDocto = 45) Then
        
                If InStr("T-D-C-F", aDoc(X).Status) = 0 Then Exit Function
            End If
        End If
    Next

    ExisteDoctoPendenteNaoTransmitido = False
    
End Function
Private Function DiferencaParaAjuste(ByVal lVinculo As Long, ByRef cTotCreditos As Currency, ByRef cTotDebitos As Currency)

Dim X As Integer

    For X = 0 To LstDocto.ListCount - 1
        If InStr("T*1", aDoc(X + 1).Status) <> 0 Then
        
'            If aDoc(X + 1).EstornoDocto = False Then
            If aDoc(X + 1).Vinculo = lVinculo And (aDoc(X + 1).EstornoDocto = False) Then
                Select Case aDoc(X + 1).TipoDocto
                    Case 4, 5, 6, 7, 41
                        cTotCreditos = cTotCreditos + aDoc(X + 1).Valor
                    Case 32, 34, 42, 44 'Soma-se como Débito para Contra partida
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 2, 3, 37      ' Depositos e OCT
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 8 To 31, 35, 36, 40   'Pagtos
                        cTotDebitos = cTotDebitos + aDoc(X + 1).Valor
                    Case 33, 38, 43, 45 'Soma-se como Crédito para Contra partida
                        cTotCreditos = cTotCreditos + aDoc(X + 1).Valor
                End Select
            End If
        End If
    Next

End Function
Private Sub AgenciaContaSomenteDeposito(ByVal lVinculo As Long, ByRef iAgenciaAjuste As Integer, ByRef lContaAjuste As Long)

' Identificar Agencia e Conta do Depósito somente se todos cheques forem do tipo (7-Chq.Depósito)
Dim X As Integer
Dim bChqDeposito As Boolean
Dim bChqNaoDeposito As Boolean

iAgenciaAjuste = 0
lContaAjuste = 0
bChqNaoDeposito = False

    For X = 0 To LstDocto.ListCount - 1
        If InStr("T*1", aDoc(X + 1).Status) <> 0 Then
        
            If aDoc(X + 1).Vinculo = lVinculo And (aDoc(X + 1).EstornoDocto = False) Then
                'Verifica se vínculo somente com depósito
                If aDoc(X + 1).TipoDocto = 37 Or aDoc(X + 1).TipoDocto = 2 Or aDoc(X + 1).TipoDocto = 3 Then
                    iAgenciaAjuste = aDoc(X + 1).DepositoAgencia
                    lContaAjuste = aDoc(X + 1).DepositoConta
                End If
                
                If aDoc(X + 1).TipoDocto = 5 Or aDoc(X + 1).TipoDocto = 6 Then
                    bChqNaoDeposito = True
                    Exit For
                End If
                
                
                If Not (aDoc(X + 1).TipoDocto = 39 Or aDoc(X + 1).TipoDocto = 37 Or _
                    aDoc(X + 1).TipoDocto = 2 Or aDoc(X + 1).TipoDocto = 3 Or _
                    aDoc(X + 1).TipoDocto = 5 Or aDoc(X + 1).TipoDocto = 6 Or aDoc(X + 1).TipoDocto = 7) Then
                    iAgenciaAjuste = 0
                    lContaAjuste = 0
                    Exit For
                End If
            End If
        End If
    Next

    If bChqNaoDeposito Then
        iAgenciaAjuste = 0
        lContaAjuste = 0
    End If
    
End Sub
