VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmShow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDI - Caixa Robô"
   ClientHeight    =   6015
   ClientLeft      =   1815
   ClientTop       =   1320
   ClientWidth     =   5940
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   7080
      TabIndex        =   31
      Top             =   870
      Width           =   3705
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   45
         TabIndex        =   34
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Terminal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   45
         TabIndex        =   33
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Estação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   45
         TabIndex        =   32
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   5910
      Begin VB.CheckBox CheckRecepcao 
         Alignment       =   1  'Right Justify
         Caption         =   "Recepcionar IK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   4410
         TabIndex        =   30
         Top             =   3600
         Width           =   1425
      End
      Begin VB.CommandButton cmdFechaCaixa 
         Caption         =   "&FECHA CAIXA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2010
         Picture         =   "frmShow.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   3180
         Width           =   1995
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         ScaleHeight     =   270
         ScaleWidth      =   4830
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2655
         Width           =   4890
         Begin MSComctlLib.ProgressBar ProgressDelayPesquisa 
            Height          =   315
            Left            =   0
            TabIndex        =   25
            Top             =   -30
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Min             =   1
            Max             =   120
         End
         Begin VB.Label LabelStatus 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   70
            LinkTimeout     =   0
            TabIndex        =   7
            Top             =   20
            Width           =   75
         End
         Begin VB.Label LabelCabecalhoStatus 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   45
            TabIndex        =   4
            Top             =   30
            Width           =   75
         End
      End
      Begin VB.Timer TimerComunica 
         Left            =   120
         Top             =   3150
      End
      Begin VB.PictureBox Picture4 
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -60
         ScaleHeight     =   2565
         ScaleWidth      =   5955
         TabIndex        =   26
         Top             =   450
         Width           =   6015
         Begin VB.Frame Frame5 
            Caption         =   "Multi-Agência:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            Left            =   2925
            TabIndex        =   42
            Top             =   0
            Width           =   2865
            Begin VB.Label LabelMultiAgencia 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "9999"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   39.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   930
               Left            =   690
               TabIndex        =   43
               Top             =   720
               Width           =   1740
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Caixa:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   135
            TabIndex        =   38
            Top             =   700
            Width           =   2700
            Begin VB.Label LabelTerminal 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "000"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   17.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Left            =   1050
               TabIndex        =   40
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Estação:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   135
            TabIndex        =   37
            Top             =   1400
            Width           =   2700
            Begin VB.Label LabelEstacao 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "000"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   17.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Left            =   1020
               TabIndex        =   41
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   135
            TabIndex        =   36
            Top             =   0
            Width           =   2700
            Begin VB.Label LabelDataProcessamento 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "99/99/9999"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   360
               Left            =   690
               TabIndex        =   39
               Top             =   300
               Width           =   1605
            End
         End
         Begin VB.Line Line1 
            X1              =   2880
            X2              =   2880
            Y1              =   0
            Y2              =   2200
         End
      End
      Begin VB.Label LabelTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "60s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   5130
         TabIndex        =   6
         Top             =   3090
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "[ F10 ] Encerra."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   75
         TabIndex        =   29
         Top             =   3630
         Width           =   1035
      End
      Begin VB.Label LabelCaixa 
         AutoSize        =   -1  'True
         Caption         =   "Caixa em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   630
         TabIndex        =   5
         Top             =   135
         Width           =   1260
      End
      Begin VB.Label LabelInstrucao 
         BackStyle       =   0  'Transparent
         Caption         =   "Processamento."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   1965
         TabIndex        =   2
         Top             =   135
         Width           =   2430
      End
   End
   Begin VB.PictureBox PictureWindowProgress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   20
      ScaleHeight     =   2100
      ScaleWidth      =   5880
      TabIndex        =   8
      Top             =   3870
      Width           =   5910
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   5775
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            ScaleHeight     =   345
            ScaleWidth      =   3420
            TabIndex        =   22
            Top             =   660
            Width           =   3450
            Begin VB.Label LabelProgressEstor 
               BackColor       =   &H000000C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   10
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            ScaleHeight     =   345
            ScaleWidth      =   3420
            TabIndex        =   15
            Top             =   240
            Width           =   3450
            Begin VB.Label LabelProgressTrans 
               BackColor       =   &H00C00000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   10
            End
         End
         Begin VB.Label LabelPercentEstor 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5145
            TabIndex        =   24
            Top             =   690
            Width           =   585
         End
         Begin VB.Label LabelPercentTrans 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5145
            TabIndex        =   21
            Top             =   300
            Width           =   585
         End
         Begin VB.Label LabelT 
            AutoSize        =   -1  'True
            Caption         =   "Transmissão:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   130
            TabIndex        =   20
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label LabelE 
            AutoSize        =   -1  'True
            Caption         =   "Estorno:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   130
            TabIndex        =   19
            Top             =   750
            Width           =   720
         End
         Begin VB.Label LabelTransmitidos 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1275
            TabIndex        =   18
            Top             =   330
            Width           =   330
         End
         Begin VB.Label LabelEstornados 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1275
            TabIndex        =   17
            Top             =   750
            Width           =   330
         End
      End
      Begin VB.Label LabelA 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Tipo Docto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   100
         TabIndex        =   28
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label LabelTipoDocto 
         AutoSize        =   -1  'True
         Caption         =   "Desconhecido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         TabIndex        =   27
         Top             =   1860
         Width           =   1050
      End
      Begin VB.Label LabelTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fora de Processamento"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -45
         TabIndex        =   13
         Top             =   0
         Width           =   5925
      End
      Begin VB.Label LabelD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Doctos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4185
         TabIndex        =   12
         Top             =   1635
         Width           =   675
      End
      Begin VB.Label LabelQtde 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5475
         TabIndex        =   11
         Top             =   1635
         Width           =   330
      End
      Begin VB.Label LabelProcessados 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1350
         TabIndex        =   10
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label LabelP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Processados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   100
         TabIndex        =   9
         Top             =   1620
         Width           =   1185
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Multi Agência"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6510
      TabIndex        =   35
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdFechaCaixa_Click()
   
   'fecha caixa aberto
    frmShow.ProgressDelayPesquisa.Tag = "FIM"

    If GetSetting("Robo", "Capa", "Corrente", 0) <> 0 Then
        DeleteSetting appname:="Robo", section:="Capa"
    End If
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        frmShow.ProgressDelayPesquisa.Tag = "OK"
    End If
    
    If KeyCode = vbKeyF10 Then
        DoEvents
        Me.cmdFechaCaixa.Visible = True
        Me.cmdFechaCaixa.Caption = "Aguarde ..."
        Call cmdFechaCaixa_Click
    End If
            
End Sub
Private Sub Form_Load()
    Dim Arquivo As String
    
    LabelTerminal.Caption = Caixa.Caixa
    LabelEstacao.Caption = Caixa.Estacao
    LabelDataProcessamento.Caption = Left(Parametros.DataServer, 2) & "/" & Mid(Parametros.DataServer, 3, 2) & "/20" & Right(Parametros.DataServer, 2)
    LabelMultiAgencia.Caption = Parametros.AgenciaCentral
    
   'Tempo Máximo em segundos do progress pesquisa
    If AntIniOpcoes.InterFixo Then
        LabelTime = "Fixo: " & AntIniOpcoes.InterValo & "s"
        ProgressDelayPesquisa.Max = AntIniOpcoes.InterValo * 2
    Else
        LabelTime.Caption = "Crescente em " & Str(AntIniOpcoes.InterValo) & " Até: 60s"
        ProgressDelayPesquisa.Max = 60 * 2
    End If

    If Dir(App.Path & "\Logs", vbDirectory) = "" Then
        MkDir App.Path & "\Logs"
    End If
    
    Arquivo = App.Path & "\Logs\EST" & Trim(Caixa.Estacao) & Parametros.AgenciaCentral & "_" & Format(Now, "ddmm") & ".TXT"
   
    Open Arquivo For Append As #20
    
End Sub
Private Sub TimerComunica_Timer()
    Static InterValo
    
    Dim i       As Integer
    Dim Passo   As Integer
        
    If AntIniOpcoes.InterFixo Then
        InterValo = AntIniOpcoes.InterValo * 2
        ProgressDelayPesquisa.Value = 1
        Passo = 1
    Else
                  
         InterValo = InterValo + 2
         Passo = AntIniOpcoes.InterValo
         
        'Se barra completa volta (1) para continuar na espera de 1 min
         If Val(InterValo) * 2 >= ProgressDelayPesquisa.Max Then
             ProgressDelayPesquisa.Value = 1
             InterValo = ProgressDelayPesquisa.Max - 1
         End If
    
    End If
    
    If frmShow.ProgressDelayPesquisa.Tag = "NO" Then
        LabelInstrucao.Caption = "STAND BY ..."
                       
        For i = 1 To Val(InterValo) - 1 Step Passo
        
            If frmShow.ProgressDelayPesquisa.Tag = "OK" Or frmShow.ProgressDelayPesquisa.Tag = "FIM" Then
               'Usuario teclou F12 p/ zerar contador ou Clicou Fechar Caixa
                Exit For
            End If
            
            Espera (0.5)
            ProgressDelayPesquisa.Value = i 'ProgressDelayPesquisa.Value + 1
            
            DoEvents
        Next i
            
        LabelInstrucao.Caption = "PROCESSAMENTO ..."
        
    End If
    
   'Finalizar o Sistema
    If frmShow.ProgressDelayPesquisa.Tag = "FIM" Then
        cmdFechaCaixa.Caption = "Encerrando ..."
        Espera (0.5)
        LogFechamentoCaixa ("T")
        Close #20
        Unload Me
        Exit Sub
    End If
    
    ProgressDelayPesquisa.Value = 1
    frmShow.ProgressDelayPesquisa.Tag = "NO"
    
    ProgressDelayPesquisa.Visible = False
    TimerComunica.Enabled = False
    
    Do 'Se encontrou capa refaz pesquisa sem delay
    Loop Until Not Comunica()
    
    ProgressDelayPesquisa.Visible = True
    TimerComunica.Enabled = True
    
End Sub
