VERSION 5.00
Object = "{ED123F48-E23F-11D4-B08D-00600899AB13}#1.0#0"; "UbbEdit.ocx"
Begin VB.Form DataBoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Captura - Data Boa"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   444
      Left            =   36
      ScaleHeight     =   390
      ScaleWidth      =   8805
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   36
      Width           =   8868
      Begin VB.Image ImgScanner 
         Height          =   480
         Left            =   180
         Picture         =   "DataBoa.frx":0000
         Top             =   0
         Width           =   480
      End
      Begin VB.Label LblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques Data Boa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   792
         TabIndex        =   29
         Top             =   36
         Width           =   2196
      End
      Begin VB.Label LblEsc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ESC] - Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   4860
         TabIndex        =   28
         Top             =   108
         Visible         =   0   'False
         Width           =   1344
      End
      Begin VB.Image ImgCheque 
         Height          =   480
         Left            =   210
         Picture         =   "DataBoa.frx":030A
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame FrameDados 
      Height          =   1284
      Left            =   792
      TabIndex        =   19
      ToolTipText     =   "Dados do cheque baixado"
      Top             =   1488
      Width           =   3648
      Begin VB.TextBox TxtCarteira 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Carteira"
         Top             =   492
         Width           =   2040
      End
      Begin VB.TextBox TxtValor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Valor"
         Top             =   876
         Width           =   2040
      End
      Begin VB.TextBox TxtBordero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   23
         Text            =   "Bordero"
         Top             =   144
         Width           =   2040
      End
      Begin VB.Label LabelValor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   84
         TabIndex        =   22
         Top             =   912
         Width           =   1308
      End
      Begin VB.Label LabelCarteira 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carteira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   84
         TabIndex        =   21
         Top             =   528
         Width           =   1308
      End
      Begin VB.Label LabelBordero 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borderô:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   84
         TabIndex        =   20
         Top             =   144
         Width           =   1308
      End
   End
   Begin VB.Frame FrameBotoes 
      Height          =   1095
      Left            =   4728
      TabIndex        =   15
      Top             =   1644
      Width           =   3360
      Begin VB.CommandButton CmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   780
         Left            =   1690
         Picture         =   "DataBoa.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Confirma Digitação."
         Top             =   200
         Width           =   792
      End
      Begin VB.CommandButton CmdCMC7 
         Caption         =   "CMC7"
         Height          =   780
         Left            =   120
         Picture         =   "DataBoa.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Digitação do CMC7"
         Top             =   200
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdLinha1 
         Caption         =   "Linha 1"
         Height          =   780
         Left            =   900
         Picture         =   "DataBoa.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Digitação da 1a. Linha do cheque."
         Top             =   200
         Width           =   780
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   780
         Left            =   2500
         Picture         =   "DataBoa.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Sai da tela - DataBoa"
         Top             =   200
         Width           =   780
      End
      Begin VB.CommandButton CmdProximo 
         Caption         =   "&Próximo Cheque"
         Height          =   780
         Left            =   1690
         Picture         =   "DataBoa.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Leitura de próximo docto."
         Top             =   200
         Width           =   780
      End
      Begin VB.CommandButton CmdScanner 
         Caption         =   "&Leitora"
         Height          =   780
         Left            =   108
         Picture         =   "DataBoa.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Inicia Leitura de (CMC7) c/  Scanner."
         Top             =   200
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin UbbEdt.UbbEdit CMC7_Campo1 
      Height          =   675
      Left            =   795
      TabIndex        =   0
      ToolTipText     =   "1o. Campo CMC7."
      Top             =   855
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1191
      TextColor       =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   11
      Title           =   "Campo 1"
   End
   Begin UbbEdt.UbbEdit CMC7_Campo2 
      Height          =   675
      Left            =   2055
      TabIndex        =   1
      ToolTipText     =   "2o. Campo CMC7."
      Top             =   855
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   12
      TextMaxNumChars =   10
      Title           =   "Campo 2"
   End
   Begin UbbEdt.UbbEdit CMC7_Campo3 
      Height          =   675
      Left            =   3450
      TabIndex        =   2
      ToolTipText     =   "3o. Campo CMC7."
      Top             =   855
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   13
      TextMaxNumChars =   12
      Title           =   "Campo 3"
   End
   Begin UbbEdt.UbbEdit Linha1_Comp 
      Height          =   675
      Left            =   810
      TabIndex        =   3
      Top             =   855
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   14
      TextMaxNumChars =   3
      Title           =   "Comp."
   End
   Begin UbbEdt.UbbEdit Linha1_Bco 
      Height          =   675
      Left            =   1500
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   3
      TextMaxNumChars =   3
      Title           =   "Banco"
   End
   Begin UbbEdt.UbbEdit Linha1_Ag 
      Height          =   675
      Left            =   2190
      TabIndex        =   5
      Top             =   855
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   4
      TextMaxNumChars =   4
      Title           =   "Agência"
   End
   Begin UbbEdt.UbbEdit Linha1_C1 
      Height          =   675
      Left            =   3030
      TabIndex        =   6
      Top             =   855
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   8
      TextMaxNumChars =   1
      Title           =   "C1"
   End
   Begin UbbEdt.UbbEdit Linha1_Conta 
      Height          =   675
      Left            =   3495
      TabIndex        =   7
      Top             =   855
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   15
      TextMaxNumChars =   10
      Title           =   "Conta"
   End
   Begin UbbEdt.UbbEdit Linha1_C2 
      Height          =   675
      Left            =   4950
      TabIndex        =   8
      Top             =   855
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   9
      TextMaxNumChars =   1
      Title           =   "C2"
   End
   Begin UbbEdt.UbbEdit Linha1_Cheque 
      Height          =   675
      Left            =   5430
      TabIndex        =   9
      Top             =   855
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   17
      TextMaxNumChars =   6
      Title           =   "Cheque"
   End
   Begin UbbEdt.UbbEdit Linha1_C3 
      Height          =   675
      Left            =   6450
      TabIndex        =   10
      Top             =   855
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   10
      TextMaxNumChars =   1
      Title           =   "C3"
   End
   Begin UbbEdt.UbbEdit Linha1_Tipo 
      Height          =   675
      Left            =   6900
      TabIndex        =   11
      Top             =   855
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1191
      TextColor       =   12582912
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextMaxNumChars =   1
      Title           =   "Tipo"
   End
   Begin UbbEdt.UBBValid UBBValid1 
      Left            =   8460
      Top             =   2190
      _ExtentX        =   794
      _ExtentY        =   820
      Banco           =   409
      Campo1          =   "CMC7_Campo1"
      Campo2          =   "CMC7_Campo2"
      Campo3          =   "CMC7_Campo3"
      Campo4          =   "Linha1_Comp"
      Campo5          =   "Linha1_Bco"
      Campo6          =   "Linha1_Ag"
      Campo7          =   "Linha1_C1"
      Campo8          =   "Linha1_Conta"
      Campo9          =   "Linha1_C2"
      Campo10         =   "Linha1_Cheque"
      Campo11         =   "Linha1_C3"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CMC7-Digitação."
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
      Left            =   792
      TabIndex        =   14
      Top             =   540
      Width           =   1728
   End
End
Attribute VB_Name = "DataBoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const L100Info = "CMC7 - Captura, Scanner (L100-Ativa)."
 Const LA93Info = "CMC7 - Captura, Scanner (LA93-Ativa)."
 Const NuloInfo = "CMC7 - Digitação."
 Const LenDifCmd_Fram = 50
 Const MsgDataboa = "Cheques Data Boa"
 
 Dim Linha1              As Boolean
 Dim TeclouEsc           As Boolean
 Dim vrCMC7              As String * 30
 Private Sub Cmc7_campo1_GotFocus()
'''''''''''''''''''''''''''''''''
'* Habilita/Desabilita Leitura *'
'''''''''''''''''''''''''''''''''
On Error GoTo Erro:
Dim Ret As enumRetornoLeitura
TeclouEsc = False

If Principal.Scanner.Scanner = eL100 Then
    If CBool(Principal.Scanner.Tag) And _
       Principal.Scanner.HABILITADO Then
RetryL100:
        Msg True
        Ret = Principal.Scanner.Le()
        Msg False
        
        If Ret = eLeituraOK Then
        
            CMC7_Campo1.Text = Principal.Scanner.CMC7_Campo1
            CMC7_Campo2.Text = Principal.Scanner.CMC7_Campo2
            CMC7_Campo3.Text = Principal.Scanner.CMC7_Campo3
            
            CmdConfirmar.SetFocus
                        
        ElseIf (Ret = eLeituraEsc And TeclouEsc) Or Ret = eTimeOut Then
            CmdCMC7.SetFocus
        ElseIf (Ret = eLeituraEsc And Not TeclouEsc) Then
            GoTo RetryL100
        ElseIf Ret = eLeituraFalha Then
            If MsgBox("Falha na Leitura... Tentar Novamente ? ", vbCritical + vbYesNo + vbApplicationModal, App.Title) = vbYes Then
                GoTo RetryL100
            Else
                CmdCMC7_Click
            End If
        End If
        
    End If
Else
    If CBool(Principal.Scanner.Tag) And _
       Principal.Scanner.HABILITADO Then
RetryLA97:
        Msg True
        Ret = Principal.Scanner.Le()
        Msg False
        
        If Ret = eLeituraOK Then
        
            CMC7_Campo1.Text = Principal.Scanner.CMC7_Campo1
            CMC7_Campo2.Text = Principal.Scanner.CMC7_Campo2
            CMC7_Campo3.Text = Principal.Scanner.CMC7_Campo3
            
            CmdConfirmar.SetFocus
            
        ElseIf Ret = eLeituraEsc And TeclouEsc Then
            CmdCMC7.SetFocus
        ElseIf (Ret = eLeituraEsc And Not TeclouEsc) Then
            GoTo RetryLA97
        ElseIf Ret = eLeituraFalha Then
            If MsgBox("Falha na Leitura... Tentar Novamente ? ", vbCritical + vbYesNo + vbApplicationModal, App.Title) = vbYes Then
                GoTo RetryLA97
            Else
                CmdCMC7_Click
            End If
        ElseIf Ret = eLeituraFim Then
            MsgBox "Alimentador do Scanner está vazio !", vbOKOnly + vbExclamation, App.Title
            cmdSair.SetFocus
        End If
    End If

End If

If Ret = eErro Then
    If Not Principal.Scanner.Erro Is Nothing Then
      Err.Raise Principal.Scanner.Erro.Number, App.Title, Principal.Scanner.Erro.Description
    End If
End If

Exit Sub

Erro:
    If Err = Principal.Scanner.Erro Then
        Call TratamentoErro("Falha no Módulo de Leitura.", Principal.Scanner.Erro, False, True)
    Else
        Call TratamentoErro("Erro no Módulo de Leitura.", Err, False, False)
    End If
    
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub cmdConfirmar_Click()
    Dim Proc_Atu        As New Custodia.Atualizar
    Dim Proc_Sel        As New Custodia.Selecionar
    Dim Ret             As Integer
 
    On Error GoTo Errostatus:
    
    If ValidaCampos Then
        Set rst = g_cMainConnection.Execute(Proc_Sel.GetChequeFusao(vrCMC7, Geral.DataProcessamento))
        
        If Not rst.EOF() Then
            If rst!fusao Then
                MsgBox "Documento já Atualizado", vbOKOnly + vbInformation
                LimpaDados True
                CMC7_Campo1.SetFocus
            Else
                TxtBordero.Text = rst!Num_Bordero
                TxtCarteira.Text = rst!CodigoCarteira
                TxtValor.Text = Format(rst!Valor, MASK_VALOR)
                
                Call g_cMainConnection.Execute(Proc_Atu.AtualizaFusao(vrCMC7, Geral.DataProcessamento), Ret, adCmdText)
                    
                If Ret = 0 Then
                    Err.Raise 998, App.Title, "Erro ao Atualizar documento - Fusão"
                    CMC7_Campo1.SetFocus
                    Exit Sub
                End If
                
                ' Nao exibe Dados do Cheque
                FrameDados.Visible = False
                CmdConfirmar.Visible = False
                ' CmdProximo.SetFocus
                CmdProximo_Click ' Elimina confirmação para o próximo cheque
                
                
            End If
        Else
                        
            If MsgBox("Documento não encontrado." & IIf(Principal.Scanner.Tag, "Continua ?", ""), IIf(Principal.Scanner.Tag, vbYesNo + vbInformation, vbOKCancel + vbInformation), App.Title) = IIf(Principal.Scanner.Tag, vbYes, vbOK) Then
                LimpaDados True
                
                If Linha1 Then
                   Linha1_Comp.SetFocus
                Else
                   CMC7_Campo1.SetFocus
                End If
            Else
                cmdSair.SetFocus
            End If
            
        End If
    End If
    
    Exit Sub
    
Errostatus:
    Call TratamentoErro("Erro ao Atualizar Fusão", Err)
    
End Sub
Private Sub CmdCMC7_Click()
    Dim Objetos  As Control

    On Error GoTo TrataErro
       
   'Se scanner conectado desabilita, senão sem efeito
    If Principal.Scanner.HABILITADO Then
        CmdCMC7.Visible = False
        CmdScanner.Visible = True
        
       'Desabilita leitura, enquanto estiver nesta seção
        Principal.Scanner.Tag = False
    End If
    
    LimpaTela Me
    LimpaDados True
    Label1.Caption = NuloInfo
    CmdConfirmar.Visible = True
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Recupera somente objetos do tipo "UbbEdit" * '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                Objetos.Visible = True
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos
    
    CMC7_Campo1.SetFocus
    
    Linha1 = False
   
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao escolher digitação por Cmc7.", Err)
    Unload Me
    
End Sub
Private Sub CmdConfirmar_GotFocus()
    cmdConfirmar_Click
End Sub

Private Sub cmdLinha1_Click()
    Dim Objetos  As Control
    
    On Error GoTo TrataErro
    
    LimpaDados True
    LimpaTela Me
    
    CmdConfirmar.Visible = True
    
    Label1.Caption = "Linha1 - Digitação"
    
   'Recupera somente objetos do tipo "UbbEdit"
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                Objetos.Visible = True
                Objetos.Text = ""
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos
    
    Linha1_Comp.SetFocus
    Linha1 = True
    
Exit Sub

TrataErro:
    Call TratamentoErro("Erro ao escolher digitação por Linha 1.", Err)
    Unload Me

End Sub
Private Sub CmdProximo_Click()
   'Se scanner for a LA93 ejeta o Cheque
    If Principal.Scanner.Scanner = eLA93 Then
        Principal.Scanner.Eject
    End If

    LimpaDados True
    CmdConfirmar.Visible = True
    If Linha1 Then
       Linha1_Comp.SetFocus
    Else
       CMC7_Campo1.SetFocus
    End If
End Sub
Private Sub cmdSair_Click()
    If Principal.Scanner.Scanner = eLA93 Then
        Principal.Scanner.Eject
    End If

    Unload Me
End Sub
Private Sub CmdScanner_Click()
   
    CmdCMC7.Visible = True
    CmdScanner.Visible = False
            
    LimpaTela Me
    LimpaDados True
    
   'Reabilita Leitura
    Principal.Scanner.Tag = True
    Label1.Caption = IIf(Principal.Scanner.Scanner = eL100, L100Info, LA93Info)
    
   'Recupera somente objetos do tipo "UbbEdit" * '
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                Objetos.Visible = True
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos
        
    CMC7_Campo1.SetFocus
    
    Linha1 = False
   
Exit Sub

TrataErro:
    Call TratamentoErro("Falha ao escolher Captura c/ Scanner.", Err)
    Unload Me

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Atualiza Flag se usuario teclou (esc) durante leitura de cheque(L100)
    If UCase(ActiveControl.Name) = "CMC7_CAMPO1" And KeyAscii = vbKeyEscape Then
        TeclouEsc = True
    End If
End Sub
Private Sub Form_Load()
    LimpaDados True
    
   'Inicializa scanner
    Call Principal.SetScanner
      
   'Principal.Scanner.Tag = False
    Label1.Caption = NuloInfo
    LblMsg.Caption = MsgDataboa
    
   'Scanner Desabilitado na primeira vez
    Principal.Scanner.Tag = False
    
    If Principal.Scanner.HABILITADO Then
        CmdScanner.Visible = True
    Else
        CmdCMC7.Visible = True
    End If
End Sub
Sub LimpaDados(Oculta As Boolean)
    LimpaTela Me
    
    FrameDados.Visible = False 'Nao Exibe os dados do Cheque
    
'    If Oculta Then
'        FrameDados.Visible = False
'    Else
'        FrameDados.Visible = True
'    End If
    
End Sub
Function ValidaCampos() As Boolean
    On Error GoTo TrataErro
    
    Dim Objetos  As Control
    Dim CMC7     As New CalculoCheque
    Dim sCMC7    As String

    ValidaCampos = False
                
    If Linha1 Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' * Valida preenchimento dos campos da Linha 1 * '
        ' * Recupera somente objetos do tipo "UbbEdit" * '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each Objetos In Me.Controls
            If TypeName(Objetos) = "UbbEdit" Then
                If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                    If Len(Trim(Objetos.Text)) = 0 Then
                        MsgBox Objetos.Title & " é obrigatório.", vbExclamation + vbOKOnly, App.Title
                        Objetos.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next Objetos
    
        ''''''''''''''''''''''''''''''''''''''''
        ' * Verifica se Linha 1 esta Correta * '
        ''''''''''''''''''''''''''''''''''''''''
        If Valida_Linha1 = False Then
            Exit Function
        End If
        
        ''''''''''''''''''''''''''''''''''
        ' * Transformar Linha1 em CMC7 * '
        ''''''''''''''''''''''''''''''''''
         CMC7.Comp = Linha1_Comp.Text
         CMC7.Banco = Linha1_Bco.Text
         CMC7.Agencia = Linha1_Ag.Text
         CMC7.Conta = Linha1_Conta.Text
         CMC7.NumeroCheque = Linha1_Cheque.Text
         CMC7.Tipificacao = Linha1_Tipo.Text
         
         '''''''''''''''''''''''''''''''''
         ' * Calcula / Retorna do CMC7 * '
         '''''''''''''''''''''''''''''''''
        If CMC7.Calcula Then
            sCMC7 = (CMC7.CMC7)
            CMC7_Campo1.Text = Mid(sCMC7, 1, 8)
            CMC7_Campo2.Text = Mid(sCMC7, 9, 10)
            CMC7_Campo3.Text = Mid(sCMC7, 19, 12)
        End If
     
    Else
        For Each Objetos In Me.Controls
            If TypeName(Objetos) = "UbbEdit" Then
                If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                    If Len(Trim(Objetos.Text)) = 0 Then
                        MsgBox Objetos.Title & " é obrigatório.", vbExclamation + vbOKOnly, App.Title
                        Objetos.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next Objetos
    End If
    
    '''''''''''''''''''''''''''''''''''''
    ' * Verifica se Cmc7 esta Correto * '
    '''''''''''''''''''''''''''''''''''''
    If Valida_CMC7 = False Then
        Exit Function
    End If
    
    vrCMC7 = (CMC7_Campo1.Text & CMC7_Campo2.Text & CMC7_Campo3.Text)
    
    ValidaCampos = True
    
    Exit Function
   
TrataErro:
    Call TratamentoErro("Erro ao validar campos.", Err)
    Unload Me
    
End Function
Sub Msg(pStatus As Boolean)
    cmdSair.Cancel = Not CBool(pStatus)
    
   'Desabilita botões durante leitura
    FrameBotoes.Enabled = Not CBool(pStatus)
    
    If pStatus Then
        Picture1.BackColor = &HC0C0FF    '&HFFFFC0
        LblMsg.Caption = "Insira Documento para Captura"
        LblEsc.Visible = True
        ImgCheque.Visible = False
        ImgScanner.Visible = True
        Screen.MousePointer = vbArrowHourglass
    Else
        Picture1.BackColor = &HC0C0C0
        LblMsg.Caption = MsgDataboa
        LblEsc.Visible = False
        ImgCheque.Visible = True
        ImgScanner.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub
Function Valida_CMC7() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                              * Verifica se CMC7 esta Correto *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos As Control
        
    Valida_CMC7 = False
    
   'Verifica Tipificação
    If Mid(CMC7_Campo2.Text, 10, 1) <> 5 Then
        MsgBox "Documento Inválido.", vbExclamation + vbOKOnly, App.Title
        If Linha1 Then
            Linha1_Comp.SetFocus
        Else
            CMC7_Campo1.Text = ""
            CMC7_Campo2.Text = ""
            CMC7_Campo3.Text = ""
            CMC7_Campo1.SetFocus
        End If
        Exit Function
    End If
            
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica cor das Fontes dos campos de CMC-7 * '
    ' * Recupera somente objetos do tipo "UbbEdit"  * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                If Objetos.TextColor = &HFF& Then
                    MsgBox "CMC7 inválido.", vbExclamation + vbOKOnly, App.Title
                    LimpaDados False
                    Exit Function
                End If
            End If
        End If
    Next Objetos

    Valida_CMC7 = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha ao validar CMC7.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Function
Function Valida_Linha1() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          * Verifica se Linha 1 esta Correta *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos As Control
        
    Valida_Linha1 = False
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica cor das Fontes dos campos de Linha1 '
    ' * Recupera somente objetos do tipo "UbbEdit"   '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                If Objetos.TextColor = &HFF& Then
                    MsgBox "Linha 1 inválida.", vbExclamation + vbOKOnly, App.Title
                    Objetos.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next Objetos

    Valida_Linha1 = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha na validação da Linha 1.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Function
Private Sub Form_Unload(Cancel As Integer)
 'Finaliza scanner
  Call Principal.DelScanner
End Sub


