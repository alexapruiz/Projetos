VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Begin VB.Form ConfFimLote 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmação Término de Captura de Lote"
   ClientHeight    =   7584
   ClientLeft      =   2664
   ClientTop       =   876
   ClientWidth     =   6828
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7584
   ScaleWidth      =   6828
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar Lote"
      Height          =   360
      Left            =   4896
      TabIndex        =   2
      Top             =   7092
      Width           =   1464
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar Captura no mesmo Lote"
      Height          =   360
      Left            =   2028
      TabIndex        =   1
      Top             =   7092
      Width           =   2796
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirmar Lote"
      Default         =   -1  'True
      Height          =   360
      Left            =   468
      TabIndex        =   0
      Top             =   7092
      Width           =   1464
   End
   Begin LeadLib.Lead Lead2 
      Height          =   2928
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3828
      Width           =   6360
      _Version        =   524288
      _ExtentX        =   11218
      _ExtentY        =   5165
      _StockProps     =   229
      BackColor       =   -2147483643
      BorderStyle     =   1
      BackErase       =   0   'False
      ScaleHeight     =   242
      ScaleWidth      =   528
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
   End
   Begin LeadLib.Lead Lead1 
      Height          =   2928
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   6360
      _Version        =   524288
      _ExtentX        =   11218
      _ExtentY        =   5165
      _StockProps     =   229
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ScaleHeight     =   242
      ScaleWidth      =   528
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imagem Frente"
      Height          =   3420
      Left            =   48
      TabIndex        =   6
      Top             =   60
      Width           =   6732
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imagem Verso"
      Height          =   3420
      Left            =   48
      TabIndex        =   7
      Top             =   3540
      Width           =   6732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique a última imagem gerada pela Vips e confirme ou não a gravação deste Lote."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   660
      Left            =   360
      TabIndex        =   5
      Top             =   6084
      Width           =   6060
   End
End
Attribute VB_Name = "ConfFimLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Resposta As Integer

Private Sub Command1_Click()
    Resposta = 1 ' Confirmar
    Me.Hide
End Sub

Private Sub Command2_Click()
    Resposta = 2 ' Continuar
    Me.Hide
End Sub

Private Sub Command3_Click()
    Resposta = 0 ' Cancelar
    Me.Hide
End Sub

Private Sub Form_Load()
    Resposta = 0 ' Cancelar
End Sub


