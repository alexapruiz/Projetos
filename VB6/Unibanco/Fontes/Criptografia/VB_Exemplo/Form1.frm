VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Teste Criptografia"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Loop"
      Height          =   435
      Left            =   2460
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Criptografa"
      Height          =   435
      Left            =   660
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração da Função de Criptografia
Private Declare Function Encripta Lib "Encripta.dll" _
        (ByVal lngIn As Long, ByVal strOut As String) As Long


Private Sub Command1_Click()
    Dim strOut As String
    
    'Eh necessario alocar espaço para o retorno
    strOut = Space$(16)
    
    'Chama a criptografia
    Encripta CLng(Val(Text1.Text)), strOut
    
    'Mostra o Retorno
    Label1.Caption = strOut
End Sub


Private Sub Command2_Click()
    Dim lngI As Long
    Dim strOut As String
    
    'Eh necessario alocar espaço para o retorno
    strOut = Space$(16)
    
    'Percorre todo range de 8 dígitos
    For lngI = 0 To 99999999
        'Chama a criptografia
        Encripta lngI, strOut
        
        'Testa se a string com menos de 16 posições ou toda zerada
        If Len(strOut) < 16 Then MsgBox "Tamanho Inválido !"
        If strOut = "0000000000000000" Then MsgBox "String Zerada !"
        
        'Atualiza a tela a cada 500 chamadas
        If lngI Mod 500 = 0 Then
            Text1.Text = CStr(lngI)
            Label1.Caption = strOut
            DoEvents
        End If
    Next
    
    'Fim
    Text1.Text = CStr(lngI)
    Label1.Caption = strOut
End Sub
