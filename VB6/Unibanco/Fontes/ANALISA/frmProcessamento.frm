VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProcessamento 
   Caption         =   " Analisador de Utilização de Scanner"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   1020
   ClientWidth     =   11880
   Icon            =   "frmProcessamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8304
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   4980
      ScaleHeight     =   2496
      ScaleWidth      =   4128
      TabIndex        =   3
      Top             =   4032
      Width           =   4152
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   372
         Left            =   1392
         TabIndex        =   59
         Top             =   2016
         Width           =   1212
      End
      Begin VB.ComboBox CboDataBase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   336
         ItemData        =   "frmProcessamento.frx":1272
         Left            =   1032
         List            =   "frmProcessamento.frx":127C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   456
         Width           =   1980
      End
      Begin ComCtl2.Animation Animation1 
         Height          =   612
         Left            =   1692
         TabIndex        =   58
         Top             =   1296
         Width           =   672
         _ExtentX        =   1185
         _ExtentY        =   1080
         _Version        =   327681
         BackColor       =   16777215
         FullWidth       =   56
         FullHeight      =   51
      End
      Begin MSMask.MaskEdBox txtDataInicial 
         Height          =   312
         Left            =   1392
         TabIndex        =   0
         Top             =   912
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         ForeColor       =   16711680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Data para Analise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4130
      End
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   10920
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   -60
      Visible         =   0   'False
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8232
      Left            =   100
      TabIndex        =   5
      Top             =   120
      Width           =   12072
      _ExtentX        =   21294
      _ExtentY        =   14520
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   441
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Utilização do Scanner"
      TabPicture(0)   =   "frmProcessamento.frx":1298
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MSFlexGrid3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSFlexGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Erros de Leitura"
      TabPicture(1)   =   "frmProcessamento.frx":12B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame7 
         Caption         =   "Erros de Leitura de CMC-7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2412
         Left            =   -74820
         TabIndex        =   56
         Top             =   360
         Width           =   5800
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   200
            Sorted          =   -1  'True
            TabIndex        =   57
            Top             =   300
            Width           =   5400
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Erros de Leitura de Código de Barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2412
         Left            =   -68880
         TabIndex        =   54
         Top             =   360
         Width           =   5800
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   200
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   300
            Width           =   5400
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5292
         Left            =   -74940
         ScaleHeight     =   5292
         ScaleWidth      =   11952
         TabIndex        =   48
         Top             =   2880
         Visible         =   0   'False
         Width           =   11952
         Begin VB.Frame Frame10 
            Caption         =   "Leitura Correta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   552
            Left            =   6024
            TabIndex        =   51
            Top             =   4560
            Width           =   5820
            Begin VB.Label lblLeituraCorreta 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "Label19"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   192
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   5592
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Leitura Scanner"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   552
            Left            =   120
            TabIndex        =   49
            Top             =   4560
            Width           =   5760
            Begin VB.Label lblLeituraScanner 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "Label19"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   192
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   5532
            End
         End
         Begin LeadLib.Lead Lead1 
            Height          =   4392
            Left            =   120
            TabIndex        =   53
            Top             =   60
            Width           =   11724
            _Version        =   524288
            _ExtentX        =   20673
            _ExtentY        =   7747
            _StockProps     =   229
            BackColor       =   16777215
            ScaleHeight     =   366
            ScaleWidth      =   977
            DataField       =   ""
            BitmapDataPath  =   ""
            AnnDataPath     =   ""
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tempos de Utilização"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Width           =   4896
         Begin VB.TextBox txtPercTempo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1224
            Width           =   960
         End
         Begin VB.TextBox txtTempoManuseio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1224
            Width           =   960
         End
         Begin VB.TextBox txtQtdeErros 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   792
            Width           =   960
         End
         Begin VB.TextBox txtTempoErro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   360
            Width           =   960
         End
         Begin VB.TextBox txtTempoCaptura 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   960
         End
         Begin VB.TextBox txtTempoConfirmacao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   792
            Width           =   960
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "% Tempo Erro"
            Height          =   192
            Left            =   2400
            TabIndex        =   44
            Top             =   1224
            Width           =   1044
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Troca de Lotes"
            Height          =   192
            Left            =   120
            TabIndex        =   43
            Top             =   1224
            Width           =   1092
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Qtde de Erros"
            Height          =   192
            Left            =   2400
            TabIndex        =   42
            Top             =   792
            Width           =   996
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Solução de Erros"
            Height          =   192
            Left            =   2400
            TabIndex        =   41
            Top             =   360
            Width           =   1248
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Confirmação"
            Height          =   192
            Left            =   120
            TabIndex        =   40
            Top             =   792
            Width           =   912
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Captura"
            Height          =   192
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   564
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Captura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1580
         Left            =   240
         TabIndex        =   25
         Top             =   420
         Width           =   2856
         Begin VB.TextBox txtQtdeDocumentosCanc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1140
            Width           =   960
         End
         Begin VB.TextBox txtQtdeDocumentos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   720
            Width           =   960
         End
         Begin VB.TextBox txtLotes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Doctos Cancelados"
            Height          =   192
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1428
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Doctos Processados"
            Height          =   192
            Left            =   120
            TabIndex        =   30
            Top             =   660
            Width           =   1524
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Qtde de Lotes"
            Height          =   192
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   1008
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtividade Média"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1580
         Left            =   3180
         TabIndex        =   20
         Top             =   420
         Width           =   2856
         Begin VB.TextBox txtProdutividade 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   300
            Width           =   960
         End
         Begin VB.TextBox txtProdutividade2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label10 
            Caption         =   "Captura + Erro"
            Height          =   192
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1392
         End
         Begin VB.Label Label1 
            Caption         =   "Captura"
            Height          =   192
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Width           =   1392
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Leitura de CMC-7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1580
         Left            =   6120
         TabIndex        =   13
         Top             =   420
         Width           =   2856
         Begin VB.TextBox txtQtdeCMC7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   300
            Width           =   960
         End
         Begin VB.TextBox txtErrosCMC7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   660
            Width           =   960
         End
         Begin VB.TextBox txtPercErrosCMC7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Reconhecido"
            Height          =   192
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total de Erros"
            Height          =   192
            Left            =   120
            TabIndex        =   18
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "% de Erros"
            Height          =   192
            Left            =   120
            TabIndex        =   17
            Top             =   1020
            Width           =   792
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Leitura de Código de Barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1580
         Left            =   9060
         TabIndex        =   6
         Top             =   420
         Width           =   2856
         Begin VB.TextBox txtQtdeCB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   300
            Width           =   960
         End
         Begin VB.TextBox txtErrosCB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   660
            Width           =   960
         End
         Begin VB.TextBox txtPercErrosCB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "% de Erros"
            Height          =   192
            Left            =   120
            TabIndex        =   12
            Top             =   1020
            Width           =   792
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total de Erros"
            Height          =   192
            Left            =   120
            TabIndex        =   11
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Reconhecido"
            Height          =   192
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   1380
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2712
         Left            =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3840
         Width           =   11652
         _ExtentX        =   20553
         _ExtentY        =   4784
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   -2147483635
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   -2147483644
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   1392
         Left            =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   6660
         Width           =   11652
         _ExtentX        =   20553
         _ExtentY        =   2455
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   -2147483635
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   -2147483644
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1590
         Left            =   5220
         TabIndex        =   47
         Top             =   2160
         Width           =   6675
         _ExtentX        =   11769
         _ExtentY        =   2815
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   -2147483635
         ForeColorFixed  =   -2147483634
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3360
      X2              =   4980
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   -60
      X2              =   4380
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuASair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Analisar &Data"
   End
   Begin VB.Menu mnuImprimir 
      Caption         =   "&Imprimir"
   End
End
Attribute VB_Name = "frmProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rdoBanco        As New ADODB.Connection
Private rdoAnalisa      As New ADODB.Connection
Private rdoTB           As ADODB.Recordset
Private rdoTB2          As ADODB.Recordset
Private sSql            As String
Private bProcessando    As Boolean
Private bCancela        As Boolean
Private nData           As Double

Private Sub cmdCancelar_Click()
   If bProcessando Then
      bCancela = True
   Else
      If Not SSTab1.Visible Then
         Unload Me
      Else
         Picture1.Visible = False
         mnuData.Enabled = True
      End If
   End If
End Sub
Private Sub Form_Load()

    Picture1.Top = (Me.Height - Picture1.Height) / 2
    Picture1.Left = (Me.Width - Picture1.Width) / 2
    Picture1.Visible = True
    Line1.X1 = 0
    Line1.X2 = Me.Width
    Line1.Y1 = 0
    Line1.Y2 = 0
    Line2.X1 = 0
    Line2.X2 = Me.Width
    Line2.Y1 = 10
    Line2.Y2 = 10

    'Inicializando os controles
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor
    CboDataBase.ListIndex = 0

    InicializaTela
End Sub
Private Function TransformaHora(pvnSegundos As Double) As String
   Dim nHora As Double
   Dim nMinuto As Double
   
   nHora = 0
   nMinuto = 0
   
   Do While 1 = 1
      If pvnSegundos >= 3600 Then
         pvnSegundos = pvnSegundos - 3600
         nHora = nHora + 1
      ElseIf pvnSegundos >= 60 Then
         pvnSegundos = pvnSegundos - 60
         nMinuto = nMinuto + 1
      Else
         Exit Do
      End If
   Loop
   
   TransformaHora = Format(nHora, "00") & ":" & Format(nMinuto, "00") & ":" & Format(pvnSegundos, "00")
End Function
Private Sub ExecutandoAnalise()

    Dim nProdutividadeOperadorCaptura    As Double
    Dim nProdutividadeOperadorErro       As Double
    Dim nQtdeOperador                    As Double
    Dim nSegundos                        As Double
    Dim nSegundosProdutividade           As Double
    Dim nSegundosProdutividade2          As Double
    Dim nSegundosCaptura                 As Double
    Dim nSegundosLoteCaptura             As Double
    Dim nSegundosLoteErro                As Double
    Dim nMediaCaptura                    As Double
    Dim nMediaCapturaErro                As Double
    Dim nQtdeLote                        As Double
    Dim nQtdeMedia                       As Double
    Dim nIndHoras                        As Double
    Dim nQtdeDoc                         As Double
    Dim nErros                           As Double
    Dim nInd                             As Double
    Dim iArq                             As Double
    Dim iArq2                            As Double
    
    Dim sErro                            As String
    Dim sOperador                        As String * 15
    Dim sArq                             As String
    Dim sDescricao                       As String * 50
    Dim sDescErro                        As String
    Dim sLinha                           As String
    Dim sArquivo                         As String
    Dim sHoraInicial                     As String
    Dim sHoraFinal                       As String
    Dim sHoraInicialErro                 As String
    Dim sHoraFinalErro                   As String
    Dim sHoraInicialOperador             As String
    Dim sHoraFinalOperador               As String
    Dim sHoraInicialTroca                As String
    Dim sHoraFinalTroca                  As String
    Dim sServidor                        As String
    Dim sBase                            As String
    Dim strUsuario                       As String
    Dim strSenha                         As String
    Dim sAgenciaCentral                  As String
    Dim DirDados                         As String
    Dim DirTrabalho                      As String
    Dim AnalisaDB                        As String
    Dim AnalisaServer                    As String
    Dim sLinhaAnterior                   As String
    Dim sDiretorioResultados             As String
    
    Dim cmd                              As ADODB.Command
    Dim cmdEstacao                       As ADODB.Command
    Dim cmdGetEstacao                    As ADODB.Command
    Dim cmdProxCodEstacao                As ADODB.Command
    Dim cmdErro                          As ADODB.Command
    Dim cmdGetErro                       As ADODB.Command
    Dim cmdGetLote                       As ADODB.Command
    Dim cmdGetTotal                      As ADODB.Command
    Dim cmdGetCaptura                    As ADODB.Command
    Dim cmdGetErroCaptura                As ADODB.Command
    Dim cmdCaptura                       As ADODB.Command
    Dim cmdProxCodCaptura                As ADODB.Command
    Dim cmdInsereLote                    As ADODB.Command
    Dim cmdInsereTotal                   As ADODB.Command
    Dim cmdErroCaptura                   As ADODB.Command
    Dim cmdCMC7_Erro                     As ADODB.Command
    Dim cmdCB_Erro                       As ADODB.Command
    
    Dim rst                              As ADODB.Recordset
    Dim rstGetEstacao                    As ADODB.Recordset
    Dim rstProxEstacao                   As ADODB.Recordset
    Dim rstGetErro                       As ADODB.Recordset
    Dim rstGetTotal                      As ADODB.Recordset
    Dim rstGetLote                       As ADODB.Recordset
    Dim rstGetErroCaptura                As ADODB.Recordset
    Dim rstGetCaptura                    As ADODB.Recordset
    Dim rstProxCodErro                   As ADODB.Recordset
    Dim rstProxCodCaptura                As ADODB.Recordset
    Dim rstCMC7_Erro                     As ADODB.Recordset
    Dim rstCB_Erro                       As ADODB.Recordset
        
    Dim ID_Estacao                       As Integer
    Dim ID_Erro                          As Integer
    Dim ID_Captura                       As Integer
    Dim Pos1                             As Integer
    Dim Pos2                             As Integer
    Dim DataProcessamento                As Long
    
    bProcessando = True

    Set cmd = New ADODB.Command
    Set cmdEstacao = New ADODB.Command
    Set cmdGetEstacao = New ADODB.Command
    Set cmdProxCodEstacao = New ADODB.Command
    Set cmdErro = New ADODB.Command
    Set cmdGetErro = New ADODB.Command
    Set cmdGetLote = New ADODB.Command
    Set cmdGetTotal = New ADODB.Command
    Set cmdGetCaptura = New ADODB.Command
    Set cmdGetErroCaptura = New ADODB.Command
    Set cmdCaptura = New ADODB.Command
    Set cmdProxCodCaptura = New ADODB.Command
    Set cmdInsereLote = New ADODB.Command
    Set cmdInsereTotal = New ADODB.Command
    Set cmdErroCaptura = New ADODB.Command
    Set cmdCMC7_Erro = New ADODB.Command
    Set cmdCB_Erro = New ADODB.Command

    On Error GoTo Err_Executando_Analise
    With rdoBanco
        sServidor = PegarOpcaoINI("Conexao", "Servidor", "MDI_NT1")
        'sBase = PegarOpcaoINI("Conexao", "DataBaseBackup", "MDI_UBB")
        sBase = CboDataBase.Text
        strUsuario = PegarOpcaoINI("Conexao", "Usuario", App.Path & "\MDI_Conexao.ini")
        strSenha = PegarOpcaoINI("Conexao", "Senha", App.Path & "\MDI_Conexao.ini")

        .ConnectionString = "driver={SQL Server};Server=" & sServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & sBase & ";provider=sqloledb"
       .Open
    End With
    
    With rdoAnalisa
        AnalisaServer = PegarOpcaoINI("Conexao", "Servidor", "MDI_NT1")
        AnalisaDB = PegarOpcaoINI("Conexao", "DataBaseDestino", "Analisa")
        .ConnectionString = "driver={SQL Server};Server=" & AnalisaServer & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & AnalisaDB & ";provider=sqloledb"
        .Open
        On Error GoTo 0

        Set cmdGetTotal.ActiveConnection = rdoAnalisa
        cmdGetTotal.CommandType = adCmdStoredProc
        cmdGetTotal.CommandText = "GetTotal"

        Set cmdGetLote.ActiveConnection = rdoAnalisa
        cmdGetLote.CommandType = adCmdStoredProc
        cmdGetLote.CommandText = "GetLote"

        Set cmdGetErro.ActiveConnection = rdoAnalisa
        cmdGetErro.CommandType = adCmdStoredProc
        cmdGetErro.CommandText = "GetErro"

        Set cmdErro.ActiveConnection = rdoAnalisa
        cmdErro.CommandType = adCmdStoredProc
        cmdErro.CommandText = "InsereErro"

        Set cmdGetCaptura.ActiveConnection = rdoAnalisa
        cmdGetCaptura.CommandType = adCmdStoredProc
        cmdGetCaptura.CommandText = "GetCaptura"

        Set cmdCaptura.ActiveConnection = rdoAnalisa
        cmdCaptura.CommandType = adCmdStoredProc
        cmdCaptura.CommandText = "InsereCaptura"

        Set cmdGetEstacao.ActiveConnection = rdoAnalisa
        cmdGetEstacao.CommandType = adCmdStoredProc
        cmdGetEstacao.CommandText = "GetEstacao"

        Set cmdProxCodEstacao.ActiveConnection = rdoAnalisa
        cmdProxCodEstacao.CommandType = adCmdStoredProc
        cmdProxCodEstacao.CommandText = "GetProxCodEstacao"

        Set cmdEstacao.ActiveConnection = rdoAnalisa
        cmdEstacao.CommandType = adCmdStoredProc
        cmdEstacao.CommandText = "InsereEstacao"

        Set cmdProxCodCaptura.ActiveConnection = rdoAnalisa
        cmdProxCodCaptura.CommandType = adCmdStoredProc
        cmdProxCodCaptura.CommandText = "GetProxCodCaptura"

        Set cmdInsereLote.ActiveConnection = rdoAnalisa
        cmdInsereLote.CommandType = adCmdStoredProc
        cmdInsereLote.CommandText = "InsereLote"

        Set cmdInsereTotal.ActiveConnection = rdoAnalisa
        cmdInsereTotal.CommandType = adCmdStoredProc
        cmdInsereTotal.CommandText = "InsereTotal"

        Set cmdCMC7_Erro.ActiveConnection = rdoAnalisa
        cmdCMC7_Erro.CommandType = adCmdStoredProc
        cmdCMC7_Erro.CommandText = "GetCMC7_Erro"

        Set cmdCB_Erro.ActiveConnection = rdoAnalisa
        cmdCB_Erro.CommandType = adCmdStoredProc
        cmdCB_Erro.CommandText = "GetCB_Erro"

        '''''''''''''''''''''''''''''''
        'Verifica se já existe na base'
        '''''''''''''''''''''''''''''''
        With cmdGetEstacao
            .Parameters(1) = Trim(sMaquina)
            Set rstGetEstacao = .Execute()
        End With
        
        If Not rstGetEstacao.EOF() Then
            With cmdGetCaptura
                .Parameters(1) = nData
                .Parameters(2) = rstGetEstacao!ID_Estacao
                Set rstGetCaptura = .Execute()
            End With
            
            '''''''''''''''''''''''''''''''''''''''''''''''
            'Se existir a captura, então recupera os dados'
            '''''''''''''''''''''''''''''''''''''''''''''''
            If Not rstGetCaptura.EOF() Then
                ''''''''''''''''''''''
                'Recuperando os dados'
                ''''''''''''''''''''''
                With rstGetCaptura
                    '''''''''''''''
                    'Frame Captura'
                    '''''''''''''''
                    txtLotes.Text = !QtdeLotesCapturados
                    txtQtdeDocumentos.Text = !QtdeLotesProcessados
                    txtQtdeDocumentosCanc.Text = !QtdeLotesCancelados
                    '''''''''''''''''''''''''''
                    'Frame Produtividade Media'
                    '''''''''''''''''''''''''''
                    txtProdutividade.Text = !ProdutividadeCaptura
                    txtProdutividade2.Text = !ProdutividadeCapturaErro
                    ''''''''''''''''''''''''''''
                    'Frame Tempos de Utilização'
                    ''''''''''''''''''''''''''''
                    txtTempoCaptura.Text = !TempoCaptura
                    txtTempoConfirmacao.Text = !TempoConfirmacao
                    txtTempoManuseio.Text = !TempoTrocaLotes
                    txtTempoErro.Text = !TempoResolucaoErros
                    txtQtdeErros.Text = !QtdeErros
                    txtPercTempo.Text = !PercTempoErro
                    '''''''''''''''''''''''
                    'Frame Leitura de CMC7'
                    '''''''''''''''''''''''
                    txtQtdeCMC7.Text = !CMC7_Reconhecido
                    txtErrosCMC7.Text = !CMC7_Erros
                    txtPercErrosCMC7.Text = !CMC7_Porcent_Erros
                    '''''''''''''''''''''''''''''''''''
                    'Frame Leitura de Codigo de Barras'
                    '''''''''''''''''''''''''''''''''''
                    txtQtdeCB.Text = !CB_Reconhecido
                    txtErrosCB.Text = !CB_Erros
                    txtPercErrosCB.Text = !CB_Porcent_Erros
                    ''''''''''''''''''''''''
                    'Preenche grid de erros'
                    ''''''''''''''''''''''''
                    With cmdGetErro
                        .Parameters(1) = nData
                        .Parameters(2) = rstGetCaptura!ID_Captura
                        Set rstGetErro = .Execute()
                    End With
                    
                    Do While Not rstGetErro.EOF()
                    
                        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                    
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 0) = rstGetErro!Cod_Erro
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 1) = rstGetErro!Descricao
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 2) = rstGetErro!Qtde
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 3) = rstGetErro!TempoParada
                    
                        rstGetErro.MoveNext
                    Loop
                    '''''''''''''''''''''''''
                    'Preenche Grid dos lotes'
                    '''''''''''''''''''''''''
                    With cmdGetLote
                        .Parameters(1) = nData
                        .Parameters(2) = rstGetCaptura!ID_Captura
                        Set rstGetLote = .Execute()
                    End With
                    Do While Not rstGetLote.EOF()
                    
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                        
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = rstGetLote!Lote
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = rstGetLote!Operador
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = rstGetLote!QtdeDocumentos
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = rstGetLote!TempoCaptura
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = rstGetLote!TempoErro
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = rstGetLote!ProdCaptura
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = rstGetLote!ProdErro
                    
                        rstGetLote.MoveNext
                    Loop
                    '''''''''''''''''''''''''
                    'Preenche Grid de totais'
                    '''''''''''''''''''''''''
                    With cmdGetTotal
                        .Parameters(1) = nData
                        .Parameters(2) = rstGetCaptura!ID_Captura
                        Set rstGetTotal = .Execute()
                    End With
                    Do While Not rstGetTotal.EOF()
                    
                        MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
                        
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 0) = rstGetTotal!Operador
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 1) = rstGetTotal!QtdeDocumentos
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 2) = rstGetTotal!TempoCaptura
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 3) = rstGetTotal!TempoErro
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 4) = rstGetTotal!ProdCaptura
                        MSFlexGrid3.TextMatrix(MSFlexGrid3.Rows - 1, 5) = rstGetTotal!ProdErro
                        
                        rstGetTotal.MoveNext
                    Loop
                    '''''''''''''''''''''''''
                    'Insere os erros de cmc7'
                    '''''''''''''''''''''''''
                    With cmdCMC7_Erro
                        .Parameters(1) = nData
                        .Parameters(2) = rstGetCaptura!ID_Captura
                        Set rstCMC7_Erro = .Execute()
                    End With
                    
                    List3.Clear
                    Do While Not rstCMC7_Erro.EOF()
                        List3.AddItem rstCMC7_Erro!Descricao
                        rstCMC7_Erro.MoveNext
                    Loop
                    '''''''''''''''''''''''
                    'Insere os erros de CB'
                    '''''''''''''''''''''''
                    With cmdCB_Erro
                        .Parameters(1) = nData
                        .Parameters(2) = rstGetCaptura!ID_Captura
                        Set rstCB_Erro = .Execute()
                    End With
                    
                    List4.Clear
                    Do While Not rstCB_Erro.EOF()
                        List4.AddItem rstCB_Erro!Descricao
                        rstCB_Erro.MoveNext
                    Loop
                    
                End With
                GoTo Fim:
            End If
        End If
        
        
    End With
        
    Set cmd.ActiveConnection = rdoBanco
    cmd.CommandText = "LerParametro"
    cmd.CommandType = adCmdStoredProc

    With cmd
        .Parameters(1) = nData
        Set rst = .Execute()
    End With

    
    If rst.EOF() Then
    
        MsgBox "Não foi possível ler os parametros do banco." & Chr(10) & _
               "Servidor = " & sServidor & Chr(10) & _
               "Base de Dados = " & sBase, vbExclamation
        rst.Close
        bProcessando = False
        rdoBanco.Close
        rdoAnalisa.Close
        Exit Sub
    
    End If
    
    
    sAgenciaCentral = rst!AgenciaCentral
    
    rst.Close
    
    ''''''''''''''''''''''''''''''''
    'Verifica se existe os arquivos'
    ''''''''''''''''''''''''''''''''
    DirDados = Dir("C:\MDI_UBB\DADOS\" & Format(nData, "00000000") & "*.TXT")
    DirTrabalho = Dir("C:\MDI_UBB\TRABALHO\DIG" & Mid(Format(nData, "00000000"), 7, 2) & Mid(Format(nData, "00000000"), 5, 2) & ".TXT")

    If DirDados = "" Or DirTrabalho = "" Then
         Beep
         MsgBox "Não há dados para serem analisados para esta data !" & Chr(10) & _
                "Os arquivos" & Chr(10) & _
                "C:\MDI_UBB\DADOS\" & Format(nData, "00000000") & "*.TXT" & Chr(10) & _
                "C:\MDI_UBB\TRABALHO\DIG" & Mid(Format(nData, "00000000"), 7, 2) & Mid(Format(nData, "00000000"), 5, 2) & ".TXT" & Chr(10) & _
                "Não existem.", vbExclamation, Caption

         txtDataInicial.Text = "__/__/____"
         rdoBanco.Close
         rdoAnalisa.Close
         bProcessando = False
         Exit Sub
    End If
       
   '''''''''''''''''''''''''''''''''''''''''''''
   'Seleciona as datas distintas da tabela capa'
   '''''''''''''''''''''''''''''''''''''''''''''
          sSql = "SELECT DataProcessamento"
   sSql = sSql & "  FROM CAPA "
   sSql = sSql & " WHERE DataProcessamento = " & nData
   
   Set rdoTB = rdoBanco.Execute(sSql)
   
   If Not rdoTB.EOF Then
      GoTo Inicio
   End If
   
   rdoTB.Close
   rdoBanco.Close
   
   bProcessando = False
   
   Beep
   MsgBox "Não existem dados para análise desta data !" & Chr(10) & _
          "Servidor = " & sServidor & Chr(10) & _
          "Base de Dados = " & sBase & Chr(10) & _
          "Tabela = CAPA", vbExclamation, Caption
   txtDataInicial.Text = "__/__/____"
   rdoAnalisa.Close
   Exit Sub
   
Inicio:
    rdoTB.Close
    
    mnuASair.Enabled = False
    
    Animation1.Visible = True
    Animation1.Open App.Path & "\FINDFILE.AVI"
    Animation1.Play
    
    iArq = FreeFile
    
    sArquivo = "C:\MDI_UBB\TRABALHO\DIG" & Mid(Format(nData, "00000000"), 7, 2) & Mid(Format(nData, "00000000"), 5, 2) & ".TXT"
    
    Open sArquivo For Input As iArq
    
    Line Input #iArq, sLinha
    
    nSegundos = 0
    nSegundosLoteCaptura = 0
    nSegundosCaptura = 0
    nIndHoras = 0
    nQtdeMedia = 0
    nMediaCaptura = 0
    
    sOperador = ""
   
   Do While 1 = 1
      sOperador = UCase(Mid(sLinha, 10, (InStr(1, sLinha, "Horario") - 11)))
      
      If InStr(1, sLinha, "Fim da captura") > 0 Then
         sHoraFinalOperador = ""
         sHoraInicialOperador = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
         sHoraFinalErro = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
         
         If sHoraInicialErro <> "" Then
            If sHoraFinalErro < sHoraInicialErro Then
               Mid(sHoraFinalErro, 1, 2) = 24 + Mid(sHoraFinalErro, 1, 2)
            End If
         
            nSegundos = ((CDbl(Mid(sHoraFinalErro, 1, 2)) * 3600) + (CDbl(Mid(sHoraFinalErro, 4, 2)) * 60) + CDbl(Mid(sHoraFinalErro, 7, 2))) - ((CDbl(Mid(sHoraInicialErro, 1, 2)) * 3600) + (CDbl(Mid(sHoraInicialErro, 4, 2)) * 60) + CDbl(Mid(sHoraInicialErro, 7, 2)))
            nSegundosLoteErro = nSegundosLoteErro + nSegundos
            txtTempoErro = CDbl(txtTempoErro) + nSegundos

            sHoraFinalErro = ""
            sHoraInicialErro = ""
            
            With MSFlexGrid2
               .Redraw = False
               
               For nInd = 0 To .Rows - 1
                  .Row = nInd
                  .Col = 0
                  If .Text = sErro Then
                     .Col = 3
                     .Text = CDbl(.Text) + nSegundos
                     Exit For
                  End If
               Next nInd
               
               .Redraw = True
            End With
         End If
      ElseIf InStr(1, sLinha, "Obtencao do numero da proxima imagem") > 0 Then
         sHoraInicial = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
         sHoraFinal = ""
      ElseIf InStr(1, sLinha, "VIPS_Captura:") > 0 Then
         sHoraFinal = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
         
         If sHoraFinal < sHoraInicial Then
            Mid(sHoraFinal, 1, 2) = 24 + Mid(sHoraFinal, 1, 2)
         End If
         
         nSegundos = ((CDbl(Mid(sHoraFinal, 1, 2)) * 3600) + (CDbl(Mid(sHoraFinal, 4, 2)) * 60) + CDbl(Mid(sHoraFinal, 7, 2))) - ((CDbl(Mid(sHoraInicial, 1, 2)) * 3600) + (CDbl(Mid(sHoraInicial, 4, 2)) * 60) + CDbl(Mid(sHoraInicial, 7, 2)))
         
         nSegundosCaptura = nSegundosCaptura + nSegundos
         nSegundosLoteCaptura = nSegundosLoteCaptura + nSegundos
         
         txtTempoCaptura = CDbl(txtTempoCaptura) + nSegundos
         
         sHoraFinal = ""

         If InStr(1, sLinha, "VIPS_Captura: 0") = 0 Then
            sHoraInicialErro = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
            txtQtdeErros = CDbl(txtQtdeErros) + 1
            
            sErro = Trim(Mid(sLinha, InStr(1, sLinha, "VIPS_Captura:") + 13, 5))
           
            Select Case sErro
               Case "-1", "-55", "-59", "-60", "-61"
                  sDescErro = "ATOLAMENTO/DESLIZAMENTO"
               Case "-3"
                  sDescErro = "COMUNICACAO"
               Case "-4", "-30"
                  sDescErro = "SINTAXE"
               Case "-50"
                  sDescErro = "TRACIONAMENTO DUPLO"
               Case "-51"
                  sDescErro = "AMASSAMENTO"
               Case "-62"
                  sDescErro = "ESCANINHO NAO ESPECIFICADO"
               Case "-101"
                  sDescErro = "ESCANINHO INVALIDO"
               Case Else
                  sDescErro = "INDEFINIDO"
            End Select
            
            With MSFlexGrid2
               .Redraw = False
               
               For nInd = 0 To .Rows - 1
                  .Row = nInd
                  .Col = 0
                  If .Text = sErro Then
                     Exit For
                  ElseIf nInd = .Rows - 1 Then
                     .Rows = .Rows + 1
                     .Row = .Rows - 1
                     .Col = 0
                     .Text = sErro
                     .Col = 1
                     .Text = sDescErro
                     .Col = 2
                     .Text = 0
                     .Col = 3
                     .Text = 0
                     nInd = .Row
                     Exit For
                  End If
               Next nInd
               
               .Col = 2
               .Text = CDbl(.Text) + 1
               
               .Redraw = True
            End With
         End If
      ElseIf InStr(1, sLinha, "Confirmacao da captura do lote") > 0 Or _
             InStr(1, sLinha, "Cancelamento da captura do lote") > 0 Then
         If InStr(1, sLinha, "Cancelamento da captura do lote") > 0 Then
            txtTempoCaptura = CDbl(txtTempoCaptura) - nSegundosCaptura
            nSegundosLoteCaptura = 0
            nSegundosLoteErro = 0
         End If
         
         nSegundosCaptura = 0
         
         sHoraFinalOperador = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
      
         If Trim(sHoraInicialOperador) <> "" Then
            nSegundos = ((Mid(sHoraFinalOperador, 1, 2) * 3600) + (Mid(sHoraFinalOperador, 4, 2) * 60) + Mid(sHoraFinalOperador, 7, 2)) - ((Mid(sHoraInicialOperador, 1, 2) * 3600) + (Mid(sHoraInicialOperador, 4, 2) * 60) + Mid(sHoraInicialOperador, 7, 2))
         End If
         
         If InStr(1, sLinhaAnterior, "Fim da captura") > 0 Then
            txtTempoConfirmacao = CDbl(txtTempoConfirmacao) + nSegundos
         End If
         
         sHoraInicialTroca = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
      ElseIf InStr(1, sLinha, "Inicio do processamento do arquivo:") > 0 Then
         iArq2 = FreeFile
         sArquivo = Mid(sLinha, InStr(1, sLinha, "Inicio do processamento do arquivo:") + 36, 21)
         
         List1.AddItem sArquivo
         
         Open "C:\MDI_UBB\DADOS\" & sArquivo For Input As iArq2
         
         nQtdeLote = 0
         
         Line Input #iArq2, sLinha
         
         Do While 1 = 1
            nQtdeLote = nQtdeLote + 1
            
            If EOF(iArq2) Then Exit Do
            
            Line Input #iArq2, sLinha
         Loop
         
         Close #iArq2
               
         If nSegundosLoteCaptura > 0 Then
            nMediaCaptura = nMediaCaptura + (((nQtdeLote / nSegundosLoteCaptura) * 3600) * nQtdeLote)
         End If
         
         If nSegundosLoteCaptura > 0 Or nSegundosLoteErro > 0 Then
            nMediaCapturaErro = nMediaCapturaErro + (((nQtdeLote / (nSegundosLoteCaptura + nSegundosLoteErro)) * 3600) * nQtdeLote)
            nQtdeMedia = nQtdeMedia + nQtdeLote
         End If
         
         With MSFlexGrid1
            .Redraw = False
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .Text = Mid(sArquivo, 9, 9)
            .Col = 1
            .Text = sOperador
            .Col = 2
            .Text = nQtdeLote
            .Col = 3
            .Text = nSegundosLoteCaptura
            .Text = TransformaHora(CDbl(.Text))
            .Col = 4
            .Text = nSegundosLoteErro
            .Text = TransformaHora(CDbl(.Text))
            .Col = 5
            If nSegundosLoteCaptura > 0 Then
               .Text = Val((nQtdeLote / nSegundosLoteCaptura) * 3600)
            End If
            .Col = 6
            If nSegundosLoteCaptura > 0 Or nSegundosLoteErro > 0 Then
               .Text = Val((nQtdeLote / (nSegundosLoteCaptura + nSegundosLoteErro)) * 3600)
            End If
            .Redraw = True
         End With
         
         With MSFlexGrid3
            .Redraw = False
            
            For nInd = 0 To .Rows - 1
               .Row = nInd
               .Col = 0
               If Trim(sOperador) = Trim(.Text) Then Exit For
               
               If nInd = .Rows - 1 Then
                  .Rows = .Rows + 1
                  .Row = .Rows - 1
                  .Col = 0
                  .Text = sOperador
                  .Col = 1
                  .Text = 0
                  .Col = 2
                  .Text = 0
                  .Col = 3
                  .Text = 0
                  .Col = 4
                  .Text = 0
                  .Col = 5
                  .Text = 0
               End If
            Next nInd
                        
            nProdutividadeOperadorCaptura = 0
            nProdutividadeOperadorErro = 0
            
            If nSegundosLoteCaptura > 0 Then nProdutividadeOperadorCaptura = (nQtdeLote / nSegundosLoteCaptura) * 3600
            If nSegundosLoteErro > 0 Or nSegundosLoteCaptura > 0 Then nProdutividadeOperadorErro = (nQtdeLote / (nSegundosLoteCaptura + nSegundosLoteErro)) * 3600
                        
            .Col = 1
            .Text = CDbl(.Text) + nQtdeLote
            .Col = 2
            .Text = CDbl(.Text) + nSegundosLoteCaptura
            .Col = 3
            .Text = CDbl(.Text) + nSegundosLoteErro
            .Col = 4
            .Text = CDbl(.Text) + (nProdutividadeOperadorCaptura * nQtdeLote)
            .Col = 5
            .Text = CDbl(.Text) + (nProdutividadeOperadorErro * nQtdeLote)
            
            .Redraw = True
         End With
         
         nSegundosLoteCaptura = 0
         nSegundosLoteErro = 0
      ElseIf InStr(1, sLinha, "Nao foi localizada a imagem do ultimo documento") > 0 Then
         sHoraInicialTroca = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
      ElseIf InStr(1, sLinha, "Inicio da captura") > 0 And _
         Len(sHoraInicialTroca) > 0 Then
         sHoraFinalTroca = Mid(sLinha, InStr(1, sLinha, "Horario:") + 9, 8)
      
         nSegundos = ((Mid(sHoraFinalTroca, 1, 2) * 3600) + (Mid(sHoraFinalTroca, 4, 2) * 60) + Mid(sHoraFinalTroca, 7, 2)) - ((Mid(sHoraInicialTroca, 1, 2) * 3600) + (Mid(sHoraInicialTroca, 4, 2) * 60) + Mid(sHoraInicialTroca, 7, 2))
         
         txtTempoManuseio = CDbl(txtTempoManuseio) + nSegundos
      End If
      
      sLinhaAnterior = sLinha
      If Not EOF(iArq) Then
         Line Input #iArq, sLinha
      Else
         Exit Do
      End If
   Loop
            
   Close #iArq
            
   With MSFlexGrid2
      For nInd = 1 To .Rows - 1
         .Row = nInd
         .Col = 3
         .Text = TransformaHora(CDbl(.Text))
      Next nInd
   End With
            
   With MSFlexGrid3
      For nInd = 1 To .Rows - 1
         .Row = nInd
         .Col = 1
         nQtdeLote = CDbl(.Text)
         .Col = 2
         .Text = TransformaHora(CDbl(.Text))
         .Col = 3
         .Text = TransformaHora(CDbl(.Text))
         .Col = 4
         .Text = Val(CDbl(.Text) / nQtdeLote)
         .Col = 5
         .Text = Val(CDbl(.Text) / nQtdeLote)
      Next nInd
   End With
            
   If nQtdeMedia > 0 Then nMediaCaptura = nMediaCaptura / nQtdeMedia
   If nQtdeMedia > 0 Then nMediaCapturaErro = nMediaCapturaErro / nQtdeMedia
   
   txtProdutividade = Val(nMediaCaptura)
   txtProdutividade2 = Val(nMediaCapturaErro)
   
   txtPercTempo = Format(((CDbl(txtTempoErro) / (CDbl(txtTempoCaptura) + CDbl(txtTempoErro))) * 100), "##0.0000")
     
   nSegundosProdutividade = CDbl(txtTempoCaptura)
   nSegundosProdutividade2 = CDbl(txtTempoCaptura) + CDbl(txtTempoErro)
     
   txtTempoCaptura = TransformaHora(CDbl(txtTempoCaptura))
   txtTempoConfirmacao = TransformaHora(CDbl(txtTempoConfirmacao))
   txtTempoManuseio = TransformaHora(CDbl(txtTempoManuseio))
   txtTempoErro = TransformaHora(CDbl(txtTempoErro))
  
   ReDim nLoteErroCMC7(1 To List1.ListCount)
   ReDim nLoteErroCB(1 To List1.ListCount)
  
   For nInd = 0 To List1.ListCount - 1
      iArq = FreeFile
      
      Open "C:\MDI_UBB\DADOS\" & Mid(List1.List(nInd), 1, 50) For Input As iArq
      
      txtLotes = CDbl(txtLotes) + 1
      nQtdeLote = 0

      Line Input #iArq, sLinha

      Do While 1 = 1
         If bCancela Then GoSub Cancelado
         
         DoEvents
      
         sSql = ""
         sSql = sSql & "Select *, '" & Mid(sLinha, 2, 50) & "' As LeituraScanner "
         sSql = sSql & "From Documento "
         sSql = sSql & "Where DataProcessamento = " & nData & " And "
         sSql = sSql & "      IdDocto           > 0 And "
         sSql = sSql & "      Frente = '" & Mid(sLinha, 65, 19) & "'"
   
         Set rdoTB = rdoBanco.Execute(sSql)
      
         If Not rdoTB.EOF Then
            If rdoTB!TipoDocto = 40 Or rdoTB!TipoDocto = 15 Then
               'MsgBox "Adivinhou"
            End If
            
            txtQtdeDocumentos = CDbl(txtQtdeDocumentos) + 1
            
            If (rdoTB!TipoDocto > 1 And rdoTB!TipoDocto < 8) Or _
               (rdoTB!TipoDocto = 1 And Len(Trim(rdoTB!Leitura)) = 14 And Mid(rdoTB!Leitura, 1, 4) = "0600") Then
               txtQtdeCMC7 = CDbl(txtQtdeCMC7) + 1
   
               If (Len(Trim(rdoTB!Leitura)) = 30 And _
                   Mid(rdoTB!Leitura, 1, 8) <> Mid(rdoTB!LeituraScanner, 2, 8) And _
                   Mid(rdoTB!Leitura, 9, 10) <> Mid(rdoTB!LeituraScanner, 11, 10) And _
                   Mid(rdoTB!Leitura, 19, 12) <> Mid(rdoTB!LeituraScanner, 22, 12)) Or _
                  (Len(Trim(rdoTB!Leitura)) = 14 And _
                   Mid(Trim(rdoTB!Leitura), 2, 9) <> Mid(rdoTB!LeituraScanner, 11, 9)) Then
   
                  txtErrosCMC7 = CDbl(txtErrosCMC7) + 1
                  sDescricao = Trim(DescricaoDocumento(rdoTB!TipoDocto))
                  List3.AddItem sDescricao & rdoTB!Frente & Space(10) & Mid(List1.List(nInd), 9, 9) & Space(10) & rdoTB!LeituraScanner & Space(10) & rdoTB!Leitura
               End If
   
               txtPercErrosCMC7 = Format((CDbl(txtErrosCMC7) / CDbl(txtQtdeCMC7)) * 100, "##0.0000")
           'documentos com 44 posições no campo leitura exceto (Darm:15 e Fgts:40)
            ElseIf (Len(Trim(rdoTB!Leitura)) = 44 And InStr(1, "15_40", rdoTB!TipoDocto) = Empty) Or _
               (rdoTB!TipoDocto = 1 And Len(Trim(rdoTB!Leitura)) = 8) Then
               txtQtdeCB = CDbl(txtQtdeCB) + 1
   
               If (rdoTB!TipoDocto = 1 And Trim(rdoTB!Leitura) <> Trim(rdoTB!LeituraScanner)) Or _
                  (rdoTB!TipoDocto <> 1 And Not IsNumeric(rdoTB!LeituraScanner)) Then
                  txtErrosCB = CDbl(txtErrosCB) + 1
                  sDescricao = Trim(DescricaoDocumento(rdoTB!TipoDocto))
                  List4.AddItem sDescricao & rdoTB!Frente & Space(10) & Mid(List1.List(nInd), 9, 9) & Space(10) & rdoTB!LeituraScanner & Space(10) & rdoTB!Leitura
               End If
   
               txtPercErrosCB = Format((CDbl(txtErrosCB) / CDbl(txtQtdeCB)) * 100, "##0.0000")
            End If
         Else
            txtQtdeDocumentosCanc = CDbl(txtQtdeDocumentosCanc) + 1
         End If
      
         rdoTB.Close
      
         If EOF(iArq) Then Exit Do
         
         Line Input #iArq, sLinha
      Loop

      Close #iArq

      nLoteErroCMC7(nInd + 1) = CDbl(txtErrosCMC7)
      nLoteErroCB(nInd + 1) = CDbl(txtErrosCB)
   Next nInd
   
   sArq = sArq & Trim(sMaquina) & "_" & Format(nData, "00000000") & ".TXT"
   
   If Dir(sArq) <> "" Then Kill sArq
   
   iArq = FreeFile
   
   DataProcessamento = Val(Right(txtDataInicial.Text, 4) + Mid(txtDataInicial.Text, 4, 2) + Left(txtDataInicial.Text, 2))
   
    '''''''''''''''''''''''''
    'Insere na base de dados'
    '''''''''''''''''''''''''
    
    rdoAnalisa.BeginTrans
    
    With cmdGetEstacao
        .Parameters(1) = Trim(sMaquina)
        Set rstGetEstacao = .Execute()
    End With

    '''''''''''''''''''''''''''''''''''
    'Se não existe a estacao, adiciona'
    '''''''''''''''''''''''''''''''''''
    If rstGetEstacao.EOF() Then
        Set rstProxEstacao = cmdProxCodEstacao.Execute()

        ID_Estacao = rstProxEstacao!ProxCodigo

        With cmdEstacao
            .Parameters(1) = ID_Estacao
            .Parameters(2) = Trim(sMaquina)
            .Execute
        End With
    Else
        ID_Estacao = rstGetEstacao!ID_Estacao
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'se não existe a captura para esta data de processamento e estação, adiciona'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With cmdGetCaptura
        .Parameters(1) = DataProcessamento
        .Parameters(2) = ID_Estacao
        Set rstGetCaptura = .Execute()
    End With

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Só adicionar a captura se não existir para esta estação'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If rstGetCaptura.EOF Then
        With cmdProxCodCaptura
            .Parameters(1) = DataProcessamento
            Set rstProxCodCaptura = .Execute()
        End With

        ID_Captura = rstProxCodCaptura!ProxCodigo

        With cmdCaptura
           .Parameters(1) = ID_Captura
           .Parameters(2) = DataProcessamento
           .Parameters(3) = ID_Estacao
           .Parameters(4) = txtTempoCaptura
           .Parameters(5) = txtTempoConfirmacao
           .Parameters(6) = txtTempoManuseio
           .Parameters(7) = txtTempoErro
           .Parameters(8) = txtQtdeErros
           .Parameters(9) = CCur(txtPercTempo)
           .Parameters(10) = Val(txtLotes)
           .Parameters(11) = Val(txtQtdeDocumentos)
           .Parameters(12) = Val(txtQtdeDocumentosCanc)
           .Parameters(13) = Val(txtProdutividade)
           .Parameters(14) = Val(txtProdutividade2)
           .Parameters(15) = Val(txtQtdeCMC7.Text)
           .Parameters(16) = Val(txtErrosCMC7.Text)
           .Parameters(17) = CCur(txtPercErrosCMC7.Text)
           .Parameters(18) = Val(txtQtdeCB.Text)
           .Parameters(19) = Val(txtErrosCB.Text)
           .Parameters(20) = CCur(txtPercErrosCB.Text)
           .Parameters(21) = ID_Erro
           .Execute
        End With
        '''''''''''''''''
        'Insere os erros'
        '''''''''''''''''
        For nInd = 1 To MSFlexGrid2.Rows - 1
            cmdErro.Parameters(1) = DataProcessamento
            cmdErro.Parameters(2) = ID_Captura
            cmdErro.Parameters(3) = MSFlexGrid2.TextMatrix(nInd, 0)
            cmdErro.Parameters(4) = MSFlexGrid2.TextMatrix(nInd, 1)
            cmdErro.Parameters(5) = MSFlexGrid2.TextMatrix(nInd, 2)
            cmdErro.Parameters(6) = MSFlexGrid2.TextMatrix(nInd, 3)
            cmdErro.Execute
        Next nInd
        '''''''''''''''''
        'Insere os Lotes'
        '''''''''''''''''
        For nInd = 1 To MSFlexGrid1.Rows - 1
            cmdInsereLote.Parameters(1) = DataProcessamento
            cmdInsereLote.Parameters(2) = ID_Captura
            cmdInsereLote.Parameters(3) = MSFlexGrid1.TextMatrix(nInd, 0)
            cmdInsereLote.Parameters(4) = Trim(MSFlexGrid1.TextMatrix(nInd, 1))
            cmdInsereLote.Parameters(5) = MSFlexGrid1.TextMatrix(nInd, 2)
            cmdInsereLote.Parameters(6) = MSFlexGrid1.TextMatrix(nInd, 3)
            cmdInsereLote.Parameters(7) = MSFlexGrid1.TextMatrix(nInd, 4)
            cmdInsereLote.Parameters(8) = Val(MSFlexGrid1.TextMatrix(nInd, 5))
            cmdInsereLote.Parameters(9) = Val(MSFlexGrid1.TextMatrix(nInd, 6))
            cmdInsereLote.Execute
        Next nInd
        ''''''''''''''''''''''''''
        'Insere o total dos Lotes'
        ''''''''''''''''''''''''''
        For nInd = 1 To MSFlexGrid3.Rows - 1
            cmdInsereTotal.Parameters(1) = DataProcessamento
            cmdInsereTotal.Parameters(2) = ID_Captura
            cmdInsereTotal.Parameters(3) = ID_Estacao
            cmdInsereTotal.Parameters(4) = Trim(MSFlexGrid3.TextMatrix(nInd, 0))
            cmdInsereTotal.Parameters(5) = MSFlexGrid3.TextMatrix(nInd, 1)
            cmdInsereTotal.Parameters(6) = MSFlexGrid3.TextMatrix(nInd, 2)
            cmdInsereTotal.Parameters(7) = MSFlexGrid3.TextMatrix(nInd, 3)
            cmdInsereTotal.Parameters(8) = MSFlexGrid3.TextMatrix(nInd, 4)
            cmdInsereTotal.Parameters(9) = MSFlexGrid3.TextMatrix(nInd, 5)
            cmdInsereTotal.Execute
        Next nInd
        
        '''''''''''''''''''''''''
        'Insere os erros de CMC7'
        '''''''''''''''''''''''''
        cmdCMC7_Erro.CommandText = "InsereCMC7_Erro"
        For nInd = 0 To List3.ListCount - 1
            cmdCMC7_Erro.Parameters(1) = nData
            cmdCMC7_Erro.Parameters(2) = ID_Captura
            cmdCMC7_Erro.Parameters(3) = List3.List(nInd)
            cmdCMC7_Erro.Execute
        Next nInd
        '''''''''''''''''''''''
        'Insere os erros de CB'
        '''''''''''''''''''''''
        cmdCB_Erro.CommandText = "InsereCB_Erro"
        For nInd = 0 To List4.ListCount - 1
            cmdCB_Erro.Parameters(1) = nData
            cmdCB_Erro.Parameters(2) = ID_Captura
            cmdCB_Erro.Parameters(3) = List4.List(nInd)
            cmdCB_Erro.Execute
        Next nInd

    End If
    
    rdoAnalisa.CommitTrans
   
    Open App.Path & "\Resultados\" & sArq For Output As iArq
    
    Print #iArq, "*****************************************************************************"
    Print #iArq, "* MDI_UBB             ANALISE DE UTILIZACAO DE SCANNER MC93                 *"
    Print #iArq, "*                                                                           *"
    Print #iArq, "* DATA DE CAPTURA - " & txtDataInicial.Text & Space(46) & "*"
    Print #iArq, "* ESTACAO         - " & Trim(sMaquina) & Space(56 - Len(Trim(sMaquina))) & "*"
    Print #iArq, "*                                                                           *"
    Print #iArq, "*****************************************************************************"
    Print #iArq, " "
    Print #iArq, "Tempo de Captura          - " & txtTempoCaptura
    Print #iArq, "Tempo de Confirmacao      - " & txtTempoConfirmacao
    Print #iArq, "Tempo de Troca de Lotes   - " & txtTempoManuseio
    Print #iArq, "Tempo de Solucao de Erros - " & txtTempoErro
    Print #iArq, "% Tempo Erro x Captura    - " & txtPercTempo
    Print #iArq, " "
    Print #iArq, "Qtde Lotes Capturados     - " & txtLotes
    Print #iArq, "Qtde Doctos Processados   - " & txtQtdeDocumentos
    Print #iArq, "Qtde Doctos Cancelados    - " & txtQtdeDocumentosCanc
    Print #iArq, "Produtividade por Hora    - " & txtProdutividade & " ( Captura )"
    Print #iArq, "Produtividade por Hora    - " & txtProdutividade2 & " ( Captura + Solução de Erros )"
    Print #iArq, " "
   
    If txtQtdeErros > 0 Then
        Print #iArq, "*************************** Utilizacao do Scanner ***************************"
        Print #iArq, "*                                                                           *"
        Print #iArq, "* Qtde de Erros Ocorridos - " & Trim(txtQtdeErros) & Space(48 - Len(Trim(txtQtdeErros))) & "*"
        Print #iArq, "* Documentos por Erro     - " & Trim(Val((CDbl(txtQtdeDocumentos) + CDbl(txtQtdeDocumentosCanc)) / CDbl(txtQtdeErros))) & Space(48 - Len(Trim(Val((CDbl(txtQtdeDocumentos) + CDbl(txtQtdeDocumentosCanc)) / CDbl(txtQtdeErros))))) & "*"
        Print #iArq, "*                                                                           *"
        Print #iArq, "* -------------- Descricao do Erro ------------- Quantidade Tempo de Parada *"
        Print #iArq, "*                                                                           *"
    
        With MSFlexGrid2
            For nInd = 1 To .Rows - 1
                .Row = nInd
                .Col = 0
                sDescErro = "[ " & .Text & Space(4 - Len(Trim(.Text))) & "]"
                .Col = 1
                sDescErro = Trim(sDescErro) & " " & .Text
                .Col = 2
                nErros = CDbl(.Text)
                .Col = 3
                Print #iArq, "* " & Trim(sDescErro) & Space(47 - Len(Trim(sDescErro))) & Space(10 - Len(Trim(Format(nErros, "###")))) & Format(nErros, "###") & Space(5) & .Text & "    *"
            Next nInd
        End With
                
        Print #iArq, "*                                                                           *"
        Print #iArq, "*****************************************************************************"
        Print #iArq, " "
    End If
         
ImprimeErrosCMC7:
   If CDbl(txtQtdeCMC7) = 0 Then GoTo ImprimeErrosCB
   
   Print #iArq, "****************************** Leitura de CMC-7 *****************************"
   Print #iArq, "*                                                                           *"
   Print #iArq, "* Total Reconhecido - " & Trim(txtQtdeCMC7) & Space(54 - Len(Trim(txtQtdeCMC7))) & "*"
   Print #iArq, "* Total de Erros    - " & Trim(txtErrosCMC7) & Space(54 - Len(Trim(txtErrosCMC7))) & "*"
   Print #iArq, "* % de Erros        - " & Trim(txtPercErrosCMC7) & Space(54 - Len(Trim(txtPercErrosCMC7))) & "*"
   Print #iArq, "*                                                                           *"
   Print #iArq, "* -------- Documentos com Leitura Incorreta -------- Quantidade Porcentagem *"
   Print #iArq, "*                                                                           *"
   
   sDescricao = Space(50)
   nQtdeDoc = 0
   
   For nInd = 0 To List3.ListCount - 1
      If Mid(List3.List(nInd), 1, 50) <> sDescricao Then
         ImprimeErroLeitura iArq, sDescricao, nQtdeDoc, List3.ListCount
         
         nQtdeDoc = 0
         sDescricao = Mid(List3.List(nInd), 1, 50)
      End If
      
      nQtdeDoc = nQtdeDoc + 1
   Next nInd
    
   ImprimeErroLeitura iArq, sDescricao, nQtdeDoc, List3.ListCount
   Print #iArq, "*                                                                           *"
   Print #iArq, "*****************************************************************************"
   Print #iArq, " "
   
ImprimeErrosCB:
   If CDbl(txtQtdeCB) = 0 Then GoTo FimImpressao
   
   Print #iArq, "************************ Leitura de Codigo de Barras ************************"
   Print #iArq, "*                                                                           *"
   Print #iArq, "* Total Reconhecido - " & Trim(txtQtdeCB) & Space(54 - Len(Trim(txtQtdeCB))) & "*"
   Print #iArq, "* Total de Erros    - " & Trim(txtErrosCB) & Space(54 - Len(Trim(txtErrosCB))) & "*"
   Print #iArq, "* % de Erros        - " & Trim(txtPercErrosCB) & Space(54 - Len(Trim(txtPercErrosCB))) & "*"
   
   Print #iArq, "*                                                                           *"
   Print #iArq, "* -------- Documentos com Leitura Incorreta -------- Quantidade Porcentagem *"
   Print #iArq, "*                                                                           *"
   
   sDescricao = Space(50)
   nQtdeDoc = 0
   
   For nInd = 0 To List4.ListCount - 1
      If Mid(List4.List(nInd), 1, 50) <> sDescricao Then
        
         ImprimeErroLeitura iArq, sDescricao, nQtdeDoc, List4.ListCount
         
         nQtdeDoc = 0
         sDescricao = Mid(List4.List(nInd), 1, 50)
      End If
      
      nQtdeDoc = nQtdeDoc + 1
   Next nInd
   
   ImprimeErroLeitura iArq, sDescricao, nQtdeDoc, List4.ListCount
   
   Print #iArq, "*                                                                           *"
   Print #iArq, "*****************************************************************************"
   
FimImpressao:
   Close iArq
   
   Me.Caption = "Analisador de Utilização de Scanner - [ " & txtDataInicial & " ]"
   
Fim:
   Picture1.Visible = False

   bProcessando = False
   
   Animation1.Stop
   Animation1.Visible = False
   
   rdoBanco.Close

   mnuASair.Enabled = True
   mnuData.Enabled = True
   'mnuGraficos.Enabled = True
   mnuImprimir.Enabled = True
   
   If List1.ListCount > 0 Then List1.ListIndex = 0
   If List1.ListCount > 0 Then List1.Selected(0) = True
   
   SSTab1.Visible = True
   
   DoEvents
   
   Set rstGetErro = Nothing
   Set rstGetCaptura = Nothing
   Set rstGetEstacao = Nothing
      
   rdoAnalisa.Close
   
   Exit Sub

Cancelado:
   Animation1.Stop
   
   If MsgBox("Deseja mesmo cancelar o processamento ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelamento de Operação") = vbNo Then
      Animation1.Play
      bCancela = False
      Return
   Else
      txtDataInicial.Text = "__/__/____"
      txtDataInicial.SetFocus
      bProcessando = False
      InicializaTela
      
      On Error Resume Next
      
      rdoTB.Close
      rdoBanco.Close
   End If
   
Exit Sub

Executando_Analise_Err:
      MsgBox Err.Description
      Exit Sub
      
Err_Executando_Analise:
      MsgBox Err.Description
      End

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bProcessando Then Cancel = 1
End Sub

Private Sub List3_Click()
   List3_DblClick
End Sub
Private Sub List3_DblClick()
   Dim nInd As Integer
   
   On Error GoTo Erro_LoadImage
   
   If bProcessando Or txtDataInicial.Text = "__/__/____" Then Exit Sub
   
   For nInd = 0 To List3.ListCount - 1
      If List3.Selected(nInd) Then
         List3.Tag = "S"
         List4.Tag = "N"
         
         Lead1.AutoSize = False
         Lead1.AutoSetRects = True
         Lead1.Load "M:\MDI_UBB\IMAGENS\" & nData & "\" & Mid(List3.List(nInd), 80, 9) & "\" & LTrim(Mid(List3.List(nInd), 50, 20)), 0, 0, 1
         
         lblLeituraScanner = Mid(List3.List(nInd), 95, 50)
         lblLeituraCorreta = Mid(List3.List(nInd), 155, 50)
         
         Picture2.Visible = True
         
         DoEvents
         
         Exit For
      End If
   Next nInd
   
   Exit Sub
Erro_LoadImage:
   
   MsgBox Error & Chr(10) & "M:\MDI_UBB\IMAGENS\" & nData & "\" & Mid(List3.List(nInd), 80, 9) & "\" & LTrim(Mid(List3.List(nInd), 50, 20)), vbExclamation
End Sub
Private Sub List3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      
      List3_DblClick
   End If
End Sub
Private Sub List4_Click()
   List4_DblClick
End Sub
Private Sub List4_DblClick()

   Dim nInd As Integer
   
   On Error GoTo Erro_LoadImage
   
   If bProcessando Or txtDataInicial.Text = "__/__/____" Then Exit Sub
   
   For nInd = 0 To List4.ListCount - 1
      If List4.Selected(nInd) Then
         List3.Tag = "N"
         List4.Tag = "S"
         
         Lead1.AutoSize = False
         Lead1.AutoSetRects = True
         Lead1.Load "M:\MDI_UBB\IMAGENS\" & nData & "\" & Mid(List4.List(nInd), 80, 9) & "\" & LTrim(Mid(List4.List(nInd), 50, 20)), 0, 0, 1
         
         lblLeituraScanner = Mid(List4.List(nInd), 95, 50)
         lblLeituraCorreta = Mid(List4.List(nInd), 155, 50)
         
         Picture2.Visible = True
         
         DoEvents
         
         Exit For
      End If
   Next nInd
   
   Exit Sub
   
Erro_LoadImage:
   
    MsgBox Error & Chr(10) & "M:\MDI_UBB\IMAGENS\" & nData & "\" & Mid(List4.List(nInd), 80, 9) & "\" & LTrim(Mid(List4.List(nInd), 50, 20)), vbExclamation
    
End Sub
Private Sub List4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      
      List4_DblClick
   End If
End Sub
Private Sub mnuASair_Click()
   Unload Me
End Sub
Private Sub mnuData_Click()
   Picture1.Visible = True
   DoEvents
   
   cmdCancelar.SetFocus
   DoEvents
   
   txtDataInicial.Text = "__/__/____"
   txtDataInicial.SetFocus
   DoEvents
End Sub
Private Sub mnuGErrosCMC7_Click()
   If txtErrosCMC7 = 0 Or txtLotes = 1 Then
      Beep
      MsgBox "Não existem dados suficientes para geração do gráfico !", vbExclamation, Caption
      Exit Sub
   End If
      
   'frmGraficoLeitura.nTipo = 1
   'frmGraficoLeitura.Show vbModal, Me
End Sub
Private Sub mnuGErrosCodigoBarras_Click()
   If txtErrosCB = 0 Or txtLotes = 1 Then
      Beep
      MsgBox "Não existem dados suficientes para a geração do gráfico !", vbExclamation, Caption
      Exit Sub
   End If
      
   'frmGraficoLeitura.nTipo = 2
   'frmGraficoLeitura.Show vbModal, Me
End Sub
Private Sub mnuGProdutividade_Click()
   If txtLotes = 1 Then
      Beep
      MsgBox "Não existem dados suficientes para geração do gráfico !", vbExclamation, Caption
      Exit Sub
   End If
      
   'frmGraficoProdutividadeDiaria.Show vbModal, Me
End Sub
Private Sub mnuImprimir_Click()

   Shell "WRITE.EXE " & App.Path & "\RESULTADOS\" & Trim(sMaquina) & "_" & Format(nData, "00000000") & ".TXT", vbMaximizedFocus
   
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If bProcessando Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If txtDataInicial.Text = "__/__/____" Then
            MsgBox "Data inválida.", vbExclamation
            Exit Sub
        End If
        KeyAscii = 0
        nData = Mid(txtDataInicial.ClipText, 5, 4) & Mid(txtDataInicial.ClipText, 3, 2) & Mid(txtDataInicial.ClipText, 1, 2)
        InicializaTela
        
        Call ExecutandoAnalise
    End If
End Sub
Private Sub InicializaTela()
   Dim nInd As Integer
   
   Me.Caption = "Analisador de Utilização de Scanner"
   
   SSTab1.Tab = 0
   SSTab1.Visible = False
   
   With MSFlexGrid1
      .Rows = 1
      .Cols = 7
      .Row = 0
      .Col = 0
      .Text = "Lote"
      .Col = 1
      .Text = "Operador"
      .Col = 2
      .Text = "Qtde Documentos"
      .Col = 3
      .Text = "Tempo de Captura"
      .Col = 4
      .Text = "Tempo de Erro"
      .Col = 5
      .Text = "Prod Captura"
      .Col = 6
      .Text = "Prod Erro"
      
      For nInd = 0 To 6
         .ColAlignment(nInd) = flexAlignCenterCenter
         .ColWidth(nInd) = .Width / 7
      Next nInd
   End With
  
   With MSFlexGrid2
      .Rows = 1
      .Cols = 4
      .Row = 0
      .Col = 0
      .Text = "Codigo"
      .Col = 1
      .Text = "Erro"
      .CellAlignment = flexAlignCenterCenter
      .Col = 2
      .Text = "Qtde"
      .Col = 3
      .Text = "Tempo de Parada"
      
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      .ColWidth(0) = .Width * 0.1
      .ColWidth(1) = .Width * 0.5
      .ColWidth(2) = .Width * 0.1
      .ColWidth(3) = .Width * 0.3
   End With
  
   With MSFlexGrid3
      .Rows = 1
      .Cols = 6
      .Col = 0
      .Text = "Operador"
      .Col = 1
      .Text = "Qtde Documentos"
      .Col = 2
      .Text = "Tempo de Captura"
      .Col = 3
      .Text = "Tempo de Erro"
      .Col = 4
      .Text = "Prod Captura"
      .Col = 5
      .Text = "Prod Erro"
      
      For nInd = 0 To 5
         .ColAlignment(nInd) = flexAlignCenterCenter
         .ColWidth(nInd) = .Width / 6
      Next nInd
   End With
  
   List1.Clear
   List3.Clear
   List4.Clear
   
   Animation1.Stop
   Animation1.Visible = False
   
   mnuASair.Enabled = True
   mnuData.Enabled = False
   'mnuGraficos.Enabled = False
   mnuImprimir.Enabled = False
   
   bCancela = False
   
   txtLotes.Text = 0
   txtQtdeDocumentos.Text = 0
   txtQtdeDocumentosCanc.Text = 0
   txtProdutividade.Text = 0
   txtProdutividade2.Text = 0
   
   txtQtdeCMC7.Text = 0
   txtErrosCMC7.Text = 0
   txtPercErrosCMC7.Text = 0
   
   txtQtdeCB.Text = 0
   txtErrosCB.Text = 0
   txtPercErrosCB.Text = 0
   
   txtTempoCaptura.Text = 0
   txtTempoConfirmacao.Text = 0
   txtTempoManuseio.Text = 0
   txtTempoErro.Text = 0
   txtPercTempo.Text = 0
   txtQtdeErros.Text = 0
End Sub
Private Function DescricaoDocumento(pvnTipoDocto As Integer) As String
   sSql = ""
   sSql = sSql & "Select * From TipoDocto "
   sSql = sSql & "Where TipoDocto = " & pvnTipoDocto
   
   'Set rdoTB2 = rdoBanco.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
   Set rdoTB2 = rdoBanco.Execute(sSql)
   
   If Not rdoTB2.EOF Then
      DescricaoDocumento = rdoTB2(1)
   Else
      DescricaoDocumento = "Docto Desconhecido"
   End If
   
   rdoTB2.Close
End Function
Private Sub ImprimeErroLeitura(ByVal pvnArq As Integer, _
                               ByVal pvsDocumento As String, _
                               ByVal pvnQtde As Double, _
                               ByVal pvnTotal As Double)
   Dim sLinha As String
   Dim nPercent As Double
   
   If pvnQtde = 0 Then Exit Sub
   
   nPercent = (pvnQtde / pvnTotal) * 100
   
   sLinha = "* " & pvsDocumento
   sLinha = sLinha & Space(11 - Len(Format(pvnQtde, "##########"))) & Format(pvnQtde, "##########")
   sLinha = sLinha & Space(11 - Len(Format(nPercent, "##0.000"))) & Format(nPercent, "#0.0000") & " *"
   
   Print #pvnArq, sLinha
End Sub
