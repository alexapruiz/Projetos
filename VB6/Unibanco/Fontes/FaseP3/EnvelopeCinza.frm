VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form EnvelopeCinza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capa de Envelope"
   ClientHeight    =   3708
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8928
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   8928
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDadosEnvelope 
      Height          =   1092
      Left            =   0
      TabIndex        =   28
      Top             =   720
      Width           =   8892
      Begin VB.TextBox Text1 
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
         Height          =   360
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   30
         Top             =   216
         Width           =   756
      End
      Begin VB.TextBox txtEnvelope 
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
         Height          =   360
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   29
         Top             =   624
         Width           =   1428
      End
      Begin VB.Label lblNomeAgencia 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "lblNomeAgencia"
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
         Height          =   252
         Left            =   3720
         TabIndex        =   33
         Top             =   264
         Width           =   5148
      End
      Begin VB.Label lblCodigoAgencia 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código da Agência :"
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
         Height          =   240
         Left            =   732
         TabIndex        =   32
         Top             =   264
         Width           =   1776
      End
      Begin VB.Label lblNumeroEnvelope 
         AutoSize        =   -1  'True
         Caption         =   "Número do Envelope:"
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
         Left            =   696
         TabIndex        =   31
         Top             =   696
         Width           =   1896
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sair"
      Height          =   720
      Left            =   8040
      Picture         =   "EnvelopeCinza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Confirmar"
      Height          =   720
      Left            =   7176
      Picture         =   "EnvelopeCinza.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Zoom -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3720
      Picture         =   "EnvelopeCinza.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zoom +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2856
      Picture         =   "EnvelopeCinza.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Frente/Verso"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6312
      Picture         =   "EnvelopeCinza.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Inverter Cor"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5448
      Picture         =   "EnvelopeCinza.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rotação"
      Height          =   720
      Left            =   4584
      Picture         =   "EnvelopeCinza.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   850
   End
   Begin VB.Frame Frame1 
      Height          =   1920
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   8892
      Begin VB.Frame Frame2 
         Caption         =   "Dados"
         Height          =   876
         Left            =   120
         TabIndex        =   11
         Top             =   984
         Width           =   5148
         Begin VB.ComboBox CboTipoConta 
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
            Height          =   336
            ItemData        =   "EnvelopeCinza.frx":1546
            Left            =   3204
            List            =   "EnvelopeCinza.frx":1548
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   444
            Width           =   1848
         End
         Begin VB.TextBox txtConta 
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
            Height          =   372
            Left            =   2136
            MaxLength       =   7
            TabIndex        =   14
            Top             =   444
            Width           =   936
         End
         Begin VB.TextBox txtIdentificado 
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
            Height          =   372
            Left            =   180
            MaxLength       =   6
            TabIndex        =   13
            Top             =   444
            Width           =   876
         End
         Begin VB.TextBox txtAgencia 
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
            Height          =   372
            Left            =   1308
            MaxLength       =   4
            TabIndex        =   12
            Top             =   444
            Width           =   672
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Conta"
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
            Height          =   240
            Left            =   3204
            TabIndex        =   19
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificado"
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
            Height          =   240
            Left            =   132
            TabIndex        =   18
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
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
            Height          =   240
            Left            =   2172
            TabIndex        =   17
            Top             =   192
            Width           =   528
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
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
            Height          =   240
            Left            =   1296
            TabIndex        =   16
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valores"
         Height          =   636
         Left            =   5352
         TabIndex        =   6
         Top             =   1224
         Width           =   3396
         Begin CURRENCYEDITLib.CurrencyEdit TxtCheques 
            Height          =   372
            Left            =   1164
            TabIndex        =   7
            Top             =   204
            Width           =   2052
            _Version        =   65537
            _ExtentX        =   3619
            _ExtentY        =   656
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin CURRENCYEDITLib.CurrencyEdit TxtTotal 
            Height          =   372
            Left            =   1044
            TabIndex        =   8
            Top             =   636
            Width           =   2076
            _Version        =   65537
            _ExtentX        =   3662
            _ExtentY        =   656
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Total"
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
            Height          =   240
            Left            =   96
            TabIndex        =   10
            Top             =   288
            Width           =   984
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Height          =   240
            Left            =   480
            TabIndex        =   9
            Top             =   720
            Width           =   444
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificação"
         Height          =   864
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5124
         Begin VB.TextBox txtCMC72 
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
            Height          =   372
            Left            =   1446
            MaxLength       =   10
            TabIndex        =   4
            Top             =   408
            Width           =   1320
         End
         Begin VB.TextBox txtCMC73 
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
            Height          =   372
            Left            =   2976
            MaxLength       =   12
            TabIndex        =   3
            Top             =   408
            Width           =   1584
         End
         Begin VB.TextBox txtCMC71 
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
            Height          =   372
            Left            =   168
            MaxLength       =   8
            TabIndex        =   2
            Top             =   408
            Width           =   1068
         End
         Begin VB.Label LblCMC7 
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "CMC-7"
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
            Height          =   192
            Left            =   180
            TabIndex        =   5
            Top             =   216
            Width           =   552
         End
      End
   End
   Begin VB.Image imgInformativo 
      Height          =   384
      Left            =   120
      Picture         =   "EnvelopeCinza.frx":154A
      Top             =   120
      Width           =   384
   End
   Begin VB.Label lblInformativo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de Envelope"
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
      Height          =   192
      Left            =   648
      TabIndex        =   27
      Top             =   276
      Width           =   1920
   End
End
Attribute VB_Name = "EnvelopeCinza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

