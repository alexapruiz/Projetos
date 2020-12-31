VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form OCT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCT"
   ClientHeight    =   3276
   ClientLeft      =   -252
   ClientTop       =   5208
   ClientWidth     =   11580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3276
   ScaleWidth      =   11580
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   750
      Left            =   9696
      Picture         =   "OCT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   750
      Left            =   10560
      Picture         =   "OCT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   750
      Left            =   6240
      Picture         =   "OCT.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   750
      Left            =   5376
      Picture         =   "OCT.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7104
      Picture         =   "OCT.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdInverteCor 
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
      Height          =   750
      Left            =   7968
      Picture         =   "OCT.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   48
      Width           =   850
   End
   Begin VB.CommandButton cmdFrenteVerso 
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
      Height          =   750
      Left            =   8832
      Picture         =   "OCT.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   48
      Width           =   852
   End
   Begin VB.Frame fraOCT 
      Height          =   2412
      Left            =   144
      TabIndex        =   0
      Top             =   816
      Width           =   11292
      Begin CURRENCYEDITLib.CurrencyEdit txtTotal 
         Height          =   360
         Left            =   9120
         TabIndex        =   30
         Top             =   1968
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
         _StockProps     =   93
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
         BackColor       =   -2147483643
      End
      Begin CURRENCYEDITLib.CurrencyEdit txtCheques 
         Height          =   360
         Left            =   9120
         TabIndex        =   29
         Top             =   1536
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
         _StockProps     =   93
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
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtDinheiro 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1152
         Width           =   1788
      End
      Begin VB.Frame fraReferencia 
         Caption         =   " Referência do Cliente "
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
         Height          =   936
         Left            =   7884
         TabIndex        =   27
         Top             =   144
         Width           =   3048
         Begin VB.TextBox txtRefConta 
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
            Left            =   1884
            MaxLength       =   7
            TabIndex        =   14
            Top             =   492
            Width           =   876
         End
         Begin VB.TextBox txtRefAgencia 
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
            Left            =   960
            MaxLength       =   4
            TabIndex        =   12
            Top             =   492
            Width           =   552
         End
         Begin VB.TextBox txtRefProd 
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
            Left            =   192
            MaxLength       =   2
            TabIndex        =   10
            Top             =   492
            Width           =   336
         End
         Begin VB.Label lblContaRef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C/C Débito"
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
            Left            =   1872
            TabIndex        =   13
            Top             =   276
            Width           =   936
         End
         Begin VB.Label lblAgenciaRef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
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
            Left            =   948
            TabIndex        =   11
            Top             =   276
            Width           =   696
         End
         Begin VB.Label lblProdutoRef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prod."
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
            Left            =   192
            TabIndex        =   9
            Top             =   276
            Width           =   456
         End
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
         Height          =   360
         Left            =   2424
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1152
         Width           =   1056
      End
      Begin VB.TextBox txtReferencia 
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
         Left            =   2436
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1560
         Width           =   3516
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
         Height          =   360
         Left            =   2424
         MaxLength       =   4
         TabIndex        =   4
         Top             =   732
         Width           =   732
      End
      Begin VB.TextBox txtOrdCredito 
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
         Left            =   2424
         MaxLength       =   15
         TabIndex        =   2
         Top             =   324
         Width           =   1812
      End
      Begin VB.Label lblNomeAgencia 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         Left            =   3312
         TabIndex        =   28
         Top             =   828
         Width           =   4524
      End
      Begin VB.Label lblRefCliente 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Referência do Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   228
         Left            =   288
         TabIndex        =   7
         Top             =   1656
         Width           =   2004
      End
      Begin VB.Label lblNumeroConta 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Número da Conta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   228
         Left            =   288
         TabIndex        =   5
         Top             =   1248
         Width           =   1740
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8028
         TabIndex        =   18
         Top             =   2076
         Width           =   588
      End
      Begin VB.Label lblDinheiro 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dinheiro"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   264
         Left            =   8028
         TabIndex        =   15
         Top             =   1248
         Width           =   912
      End
      Begin VB.Label lblCheque 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8028
         TabIndex        =   17
         Top             =   1656
         Width           =   912
      End
      Begin VB.Label lblAgenciaCred 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Agência de Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   264
         Left            =   288
         TabIndex        =   3
         Top             =   828
         Width           =   1920
      End
      Begin VB.Label lblOrdemCred 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Ord.Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   264
         Left            =   288
         TabIndex        =   1
         Top             =   420
         Width           =   2088
      End
   End
   Begin VB.Image imgInformativo 
      Height          =   384
      Left            =   528
      Picture         =   "OCT.frx":1546
      Top             =   288
      Width           =   384
   End
   Begin VB.Label lblInformativo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de OCT"
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
      Left            =   1296
      TabIndex        =   26
      Top             =   348
      Width           =   1500
   End
End
Attribute VB_Name = "OCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variavel de retorno informando se Cancelou ou Alterou
Public Alterou As Boolean

Dim bAlterar                As Boolean
Dim sPosicaoErro            As String
Private mForm               As Form

Private Type tpModulo
    qryInserirOCT           As rdoQuery
    qryChecarAgencia        As rdoQuery
    qryGetOCT               As rdoQuery
    rstModulo               As rdoResultset
    qryRemoveTipoDocumento  As rdoQuery
End Type

Private Modulo              As tpModulo
Private Function VerificarTudo() As Boolean
    
    VerificarTudo = False
        
    'Valida número ordem de credito
    If Not VerificarOrdCredito Then
        Exit Function
    End If
    
    'Valida Código da agência
    If Not AgenciaOk Then
        txtAgencia_GotFocus
        txtAgencia.SetFocus
        Exit Function
    End If

    'Valida Agencia e Conta
    If Not VerificarAgenciaConta Then
        Exit Function
    Else
        'conforme especificação do Unibanco
        If (txtAgencia = "0089") And (txtConta = "1209363") Then
            txtReferencia = ""
            Call ReferenciaCliente(True)
            
            'Valida Ref. Produto
            If Not VerificarRefProd Then
                Exit Function
            End If
            
            'Valida Ref. Agencia e Conta
            If Not VerificarAgenciaContaRef Then
                Exit Function
            End If

        Else
            Call ReferenciaCliente(False)
            txtRefProd = ""
            txtRefAgencia = ""
            txtRefConta = ""

        End If
    End If
    
    If Len(Trim(txtCheques.Text)) = 0 Or Val(Desformata_Valor(txtCheques.Text)) = 0 Then
        MsgBox "O Valor dos Cheques deve ser informado!", vbExclamation + vbOKOnly, App.Title
        txtCheques.SetFocus
        Exit Function
    End If
    
    VerificarTudo = True
    
End Function

Private Function VerificarAgenciaContaRef() As Boolean
    
    VerificarAgenciaContaRef = False
    
    If Val(txtRefAgencia) = 0 Then
        MsgBox "O preenchimento do campo Ref. Agencia é obrigatório.", vbInformation, App.Title
        txtRefAgencia_GotFocus
        txtRefAgencia.SetFocus
        Exit Function
    Else
        txtRefAgencia = Format(txtRefAgencia, "0000")
    End If
    
    If Val(txtRefConta) = 0 Then
        MsgBox "O preenchimento do campo Ref. Conta é obrigatório.", vbInformation, App.Title
        txtRefConta_GotFocus
        txtRefConta.SetFocus
        Exit Function
    Else
        If Len(Trim(txtRefConta)) < 7 Then
            txtRefConta = Format(txtRefConta, "0000000")
        End If
        If Not Modulo10(txtRefAgencia.Text & txtRefConta.Text, 11) Then
            MsgBox "Agência e Conta inválidos!", vbExclamation + vbOKOnly, App.Title
            txtRefAgencia_GotFocus
            txtRefAgencia.SetFocus
            Exit Function
        End If
    End If
    VerificarAgenciaContaRef = True
End Function

Private Function VerificarRefProd() As Boolean
    VerificarRefProd = False
    
    If Val(txtRefProd) = 0 Then
        MsgBox "O preenchimento do campo Prod. é obrigatório.", vbInformation, App.Title
        txtRefProd_GotFocus
        txtRefProd.SetFocus
        Exit Function
    Else
        If (Val(txtRefProd) < 11) Or (Val(txtRefProd) > 32) Then
            MsgBox "Este produto não pode ser aceito conforme regra da OCT.", vbInformation, App.Title
            txtRefProd = ""
            txtRefProd.SetFocus
            Exit Function
        End If
    End If
    
    VerificarRefProd = True

End Function

Private Function VerificarAgenciaConta() As Boolean
    
    VerificarAgenciaConta = False
    
    If Len(Trim(txtAgencia)) = 0 Then
        MsgBox "O preenchimento do campo Agência é obrigatório.", vbInformation, App.Title
        txtAgencia_GotFocus
        txtAgencia.SetFocus
        Exit Function
    Else
        txtAgencia = Format(txtAgencia, "0000")
    End If
    
    If Len(Trim(txtConta)) = 0 Then
        MsgBox "O preenchimento do campo Conta é obrigatório.", vbInformation, App.Title
        txtConta_GotFocus
        txtConta.SetFocus
        Exit Function
    Else
        If Len(Trim(txtConta)) < 7 Then
            txtConta = Format(txtConta, "0000000")
        End If
        
        If Not ValidaConta(Left(txtConta, 6)) Then
            MsgBox "Esta conta não pode ser aceita conforme regras de recebimento de OCT.", vbExclamation + vbOKOnly, App.Title
            txtConta_GotFocus
            txtConta.SetFocus
            Exit Function
        Else
            If Not Modulo10(txtAgencia.Text & txtConta.Text, 11) Then
                MsgBox "Agência e Conta inválidos!", vbExclamation + vbOKOnly, App.Title
                txtAgencia_GotFocus
                txtAgencia.SetFocus
                Exit Function
            End If
        End If
    End If
    
    VerificarAgenciaConta = True

End Function

Private Function VerificarOrdCredito() As Boolean

    VerificarOrdCredito = False
    
    If Len(Trim(txtOrdCredito)) = 0 Then
        MsgBox "O preenchimento do campo Número Ordem Crédito é obrigatório", vbInformation, App.Title
        txtOrdCredito_GotFocus
        txtOrdCredito.SetFocus
        Exit Function
    Else
        txtOrdCredito = Format(txtOrdCredito, "000000000000000")
        ' Calculo do Modulo 11(mesmo do envelope) ou Simplificado
        
        If Right(txtOrdCredito, 1) <> Modulo11UBB(Val(Left(txtOrdCredito, Len(txtOrdCredito) - 1))) Then
            If Right(txtOrdCredito, 1) <> Modulo11Simplificado(Val(Left(txtOrdCredito, Len(txtOrdCredito) - 1))) Then
                MsgBox "Digito verificador Ordem Crédito não confere", vbExclamation + vbOKOnly, App.Title
                txtOrdCredito_GotFocus
                txtOrdCredito.SetFocus
                Exit Function
            End If
        End If
    End If
    
    VerificarOrdCredito = True

End Function
Function ValidaConta(ByVal n_conta As String) As Boolean
'*****************************************************************************************
'* Unibanco (Marcelo) passou informação para não receber apenas contas no intervalo de   *
'* 680.000 a 689.000 - se for menor ou maior que este número, pode receber.              *
'* Leda alterou em 15/06/2000                                                            *
'*****************************************************************************************
'    'Intervalos validos (dados passados pelo UBB)
'    'de 100001 até 200000 (exceto 199990,199991 ou 199994)
'    'de 300001 até 510000, de 520001 até 680000, de 700001 até 799999.
'    If (Val(n_conta) = 199990) Or (Val(n_conta) = 199991) Or (Val(n_conta) = 199994) Then
'        ValidaConta = False
'        Exit Function
'    End If
'
'
'    If ((Val(n_conta) >= 100001) And (Val(n_conta) <= 200000)) Or _
'       ((Val(n_conta) >= 300001) And (Val(n_conta) <= 510000)) Or _
'       ((Val(n_conta) >= 520001) And (Val(n_conta) <= 680000)) Or _
'       ((Val(n_conta) >= 700001) And (Val(n_conta) <= 799999)) Then
'       ValidaConta = True
'    Else
'       ValidaConta = False
'    End If
    
    '********************************
    '* Regra passada em 15/06/2000  *
    '********************************
    If ((Val(n_conta) < 680000) Or (Val(n_conta) > 689000)) Then
       ValidaConta = True
    Else
       ValidaConta = False
    End If

End Function

Private Sub cmdConfirmar_Click()

On Error GoTo Err_cmdConfirmar

    Dim strEncripta   As String
    
    'Valida todos os campos
    If VerificarTudo Then
            
        'Força Valor de Cheque em Total
        txtTotal.Text = Format(Val(txtCheques.Text) / 100, "###,###,##0.00  ")
        
        'Inicia Transação
        Geral.Banco.BeginTrans
        
        'Verificar se o Documento pertence à outro Tipo
        If Geral.Documento.TipoDocto <> 37 And Geral.Documento.TipoDocto <> 0 Then
          With Modulo.qryRemoveTipoDocumento
            .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
            .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
            .Execute
          End With
        End If
        
        
        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(37, CStr(Val(txtConta.Text)))
        If strEncripta = "" Then GoTo Exit_cmdConfirmar
        
        sPosicaoErro = "InsOCT"
        With Modulo.qryInserirOCT
            .rdoParameters(1) = Geral.DataProcessamento                         'Data de processamento
            .rdoParameters(2) = Geral.Documento.IdDocto                         'IdDocto
            .rdoParameters(3) = Val(txtOrdCredito.Text)                         'Ordem credito
            .rdoParameters(4) = Val(txtAgencia.Text)                            'agencia credito
            .rdoParameters(5) = Val(txtConta.Text)                              'conta credito
            If (txtAgencia.Text = "0089") And (txtConta.Text = "1209363") Then
                .rdoParameters(6) = txtRefProd & txtRefAgencia & txtRefConta    'referencia
            Else
                .rdoParameters(6) = txtReferencia.Text                          'referencia
            End If
            .rdoParameters(7) = Val(txtRefProd.Text)                            'produto
            .rdoParameters(8) = Val(txtRefAgencia.Text)                         'ag.cliente
            .rdoParameters(9) = Val(txtRefConta.Text)                           'cta.cliente
            .rdoParameters(10) = Val(Desformata_Valor(txtDinheiro.Text)) / 100  'Valor em Dinheiro
            .rdoParameters(11) = Val(Desformata_Valor(txtCheques.Text)) / 100   'Valor em Cheque
            .rdoParameters(12) = Val(Desformata_Valor(txtTotal.Text)) / 100     'Valor Total
            .rdoParameters(13) = strEncripta                                    'Autenticacao digital
            .Execute
            
            If .rdoParameters(0).Value <> 0 Then GoTo Exit_cmdConfirmar
        End With
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Envia para confirmação somente se o usuario for terceiro'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If GrupoUsuario(Geral.Usuario, eG_TERCEIRO) Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Não faz nada caso não conseguiu atualizar o status do documento'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not ConfirmaAgConta(Geral.Documento.IdDocto) Then
                MsgBox "Não foi possível enviar este documento para confirmação de Agência e Conta.", vbCritical
                Exit Sub
            End If
            Geral.Documento.Status = "L"
        End If

        
        'Finaliza Transação
        Geral.Banco.CommitTrans
        
        Alterou = True
        Me.Hide
    End If
    
    Exit Sub
    
Exit_cmdConfirmar:
    Alterou = False
    Geral.Banco.RollbackTrans
    Exit Sub

Err_cmdConfirmar:
   
    Alterou = False
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível atualizar/inserir o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    Me.Hide

End Sub

Private Sub cmdFrenteVerso_Click()
    
    mForm.cmdFrenteVerso_Click
    
End Sub

Private Sub cmdInverteCor_Click()

    mForm.cmdInverteCor_Click
    
End Sub

Private Sub cmdRotacao_Click()

    mForm.cmdRotacao_Click
    
End Sub

Private Sub CmdSair_Click()
    Alterou = False
    Me.Hide
End Sub

Private Sub cmdZoomMais_Click()

    mForm.cmdZoomMais_Click
    
End Sub

Private Sub cmdZoomMenos_Click()

    mForm.cmdZoomMenos_Click
    
End Sub

Private Sub Form_Activate()
    
    'Verifica se Capa de OCT já cadastrada
    If Geral.Documento.Status = "1" Or Geral.Documento.Status = "L" Or Geral.Documento.Status = "Y" Then
        With Modulo.qryGetOCT
            .rdoParameters(0) = Geral.DataProcessamento     'Data de Processamento
            .rdoParameters(1) = Geral.Documento.IdDocto     'IdDocto
            Set Modulo.rstModulo = .OpenResultset(rdOpenStatic)
            If Not Modulo.rstModulo.EOF() Then
                txtOrdCredito.Text = Modulo.rstModulo!OrdemCredito
                txtAgencia.Text = Modulo.rstModulo!AgenciaCredito
                txtConta.Text = Modulo.rstModulo!ContaCredito
                If Modulo.rstModulo!AgenciaCredito = 89 And _
                    Modulo.rstModulo!ContaCredito = 1209363 Then
                    txtReferencia = ""
                    Call ReferenciaCliente(True)

                Else
                    txtReferencia = Modulo.rstModulo!Referencia
                    Call ReferenciaCliente(False)
                End If
                
                txtRefProd = IIf(Modulo.rstModulo!Produto = 0, "", Modulo.rstModulo!Produto)
                txtRefAgencia = IIf(Modulo.rstModulo!AgCliente = 0, "", Modulo.rstModulo!AgCliente)
                txtRefConta = IIf(Modulo.rstModulo!CtaCliente = 0, "", Modulo.rstModulo!CtaCliente)
                txtDinheiro.Text = Format(Modulo.rstModulo!Dinheiro, "###,###,##0.00  ")
                txtCheques.Text = Modulo.rstModulo!Cheque * 100
                txtTotal.Text = Format(Modulo.rstModulo!Valor * 100)

            End If
        End With
    Else
        Call ReferenciaCliente(False)
    End If
   
    txtOrdCredito_GotFocus
    txtOrdCredito.SetFocus
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      cmdZoomMais_Click
    Case vbKeySubtract
      cmdZoomMenos_Click
    Case vbKeyF10
      cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      cmdRotacao_Click
    Case vbKeyF11
      cmdFrenteVerso_Click
    Case vbKeyMultiply
        Call cmdConfirmar_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        mForm.Form_KeyUp KeyCode, Shift
  End Select
End Sub

Private Sub Form_Load()
    bAlterar = True
    
    Alterou = False
    
    'Desabilita controles
    lblDinheiro.Enabled = False
    lblTotal.Enabled = False
    txtDinheiro.ForeColor = vbBlack
    txtDinheiro.BackColor = G_ColorGray
    txtDinheiro.Locked = True
    
    txtTotal.ForeColor = vbBlack
    txtTotal.BackColor = G_ColorGray
    txtTotal.Locked = True
    
    'Query para a gravação dos dados do OCT
    Set Modulo.qryInserirOCT = Geral.Banco.CreateQuery("", "{? = call InserirOCT (?,?,?,?,?,?,?,?,?,?,?,?,?)}")

    'Query para verificação do Códig da agência
    Set Modulo.qryChecarAgencia = Geral.Banco.CreateQuery("", "{call ObtemAgencia (?)}")

    'Query para ler dados da tabela OCT
    Set Modulo.qryGetOCT = Geral.Banco.CreateQuery("", "{call GetOCT (?,?)}")
        
    Set Modulo.qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    With Modulo
        .qryInserirOCT.Close
        .qryChecarAgencia.Close
        .qryGetOCT.Close
        .qryRemoveTipoDocumento.Close
    End With

End Sub

Private Sub txtAgencia_Change()
    
    lblNomeAgencia.Caption = ""
    
End Sub

Private Sub txtAgencia_GotFocus()
    With txtAgencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAgencia = Format(txtAgencia, "0000")
        
        If Val(txtAgencia) = 0 Then
            MsgBox "O preenchimento deste campo é obrigatório.", vbInformation, App.Title
            txtAgencia_GotFocus
            txtAgencia.SetFocus
            Exit Sub

        End If
        
        If Not AgenciaOk Then
            txtAgencia_GotFocus
            txtAgencia.SetFocus
            Exit Sub
        End If
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
    
End Sub
Private Sub txtCheques_Change()
    
    Dim nparou As Integer
    Dim ntamanho As String
    
    If Not bAlterar Then
        Exit Sub
    End If
    
    bAlterar = False
    
    With txtCheques
        'Guarda posição do cursor
        nparou = .SelStart
        'Guarda tamanho texto
        ntamanho = Len(Trim(.Text))
        'Chama função formata texto
        .Text = Formata_Valor(.Text)
        .SelLength = 0
        'Calcula nova posição do cursor
        nparou = nparou + (Len(Trim(.Text)) - ntamanho)
        If nparou < 0 Then nparou = 0
        .SelStart = nparou
    End With
    
    bAlterar = True
End Sub

Private Sub txtCheques_GotFocus()

    With txtCheques
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
Private Sub txtCheques_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    
    If KeyAscii = 13 Then
        If Len(Trim(txtCheques.Text)) = 0 Or Val(Desformata_Valor(txtCheques.Text)) = 0 Then
            MsgBox "O Valor dos Cheques deve ser informado!", vbExclamation + vbOKOnly, App.Title
            txtCheques.SetFocus
            Exit Sub
        Else
            txtTotal.Text = txtCheques.Text
        End If
        
        'Finaliza digitação
        KeyAscii = 0
        cmdConfirmar_Click
    
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub
Private Sub txtConta_GotFocus()
    With txtConta
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtConta_KeyPress(KeyAscii As Integer)

    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Then
    
        'Valida Agencia e Conta
        If VerificarAgenciaConta Then
            'conforme especificação do Unibanco
            If (txtAgencia = "0089") And (txtConta = "1209363") Then
                txtReferencia = ""
                
                Call ReferenciaCliente(True)
                txtRefProd.SetFocus
            Else
                Call ReferenciaCliente(False)
                
                txtRefProd = ""
                txtRefAgencia = ""
                txtRefConta = ""
                txtReferencia.SetFocus
            End If
        End If
        
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub

Private Sub txtDinheiro_GotFocus()
    SendKeys "{TAB}"
End Sub

Private Sub txtOrdCredito_GotFocus()
    With txtOrdCredito
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOrdCredito_KeyPress(KeyAscii As Integer)
   
    InibirTeclaAlfa KeyAscii
   
    If (KeyAscii = 13) Then
        
        KeyAscii = 0
        
        'Valida Número Ordem Crédito
        If VerificarOrdCredito Then
            SendKeys "{TAB}"
        End If
   
   ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
   End If
End Sub

Private Sub txtRefAgencia_GotFocus()

    txtRefAgencia.SelStart = 0
    txtRefAgencia.SelLength = txtRefAgencia.MaxLength

End Sub

Private Sub txtRefAgencia_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Then
        If Val(txtRefAgencia) = 0 Then
            MsgBox "O preenchimento deste campo é obrigatório.", vbInformation, App.Title
            txtRefAgencia.SetFocus
        Else
            If Len(Trim(txtRefAgencia)) < 4 Then
                txtRefAgencia = Format(txtRefAgencia, "0000")
            End If
        End If
        SendKeys "{TAB}"
            
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub
Private Sub txtRefConta_GotFocus()
    With txtRefConta
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRefConta_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Then
    
        'Valida Agencia e Conta Referencia
        If Not VerificarAgenciaContaRef Then Exit Sub
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
    
End Sub
Private Sub txtReferencia_GotFocus()
    With txtReferencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    
    'Não permite digitação para evitar problema no ontime
    If KeyAscii = Asc(",") Or KeyAscii = Asc(":") Then
          KeyAscii = 0
          Exit Sub
     End If
    
    If KeyAscii = 13 Then
        txtCheques.SetFocus
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    
    ElseIf KeyAscii = vbKeyExecute Or KeyAscii = vbKeyInsert Or KeyAscii = vbKeyHelp Or KeyAscii = vbKeyPrint Then
         KeyAscii = 0
    End If

End Sub
Private Sub txtRefProd_GotFocus()
    With txtRefProd
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRefProd_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Then
    
        'Valida Ref Produto
        If Not VerificarRefProd Then Exit Sub
        
        SendKeys "{TAB}"
    
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub

Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

Private Function AgenciaOk() As Boolean

Dim rstModulo As rdoResultset

On Error GoTo Err_AgenciaOk

    AgenciaOk = False
    
    'Verifica se agência é válida
    If Len(Trim(txtAgencia.Text)) = 0 Then
        Beep
        MsgBox "A Agência de origem deve ser informada!", vbExclamation + vbOKOnly, App.Title
        GoTo Exit_AgenciaOk
    End If
    
    With Modulo.qryChecarAgencia
        .rdoParameters(0) = Val(txtAgencia.Text)
        Set rstModulo = .OpenResultset(rdOpenStatic)
    End With

    If rstModulo.RowCount > 0 Then
        lblNomeAgencia.Caption = rstModulo!agefsnoagen
    Else
        Beep
        MsgBox "Código de Agência inválido. Verifique!", vbInformation, App.Title
        GoTo Exit_AgenciaOk
    End If

    AgenciaOk = True


Exit_AgenciaOk:

    If Not (rstModulo Is Nothing) Then rstModulo.Close
    Exit Function
    
Err_AgenciaOk:
   
    'Fecha Resultset
    If Not (rstModulo Is Nothing) Then rstModulo.Close
  
    Select Case TratamentoErro("Não foi possível verificar agência.", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    Err.Clear
    GoTo Exit_AgenciaOk
    
End Function

Private Sub ReferenciaCliente(bHabilita As Boolean)

    If bHabilita Then
        txtRefProd.BackColor = vbWhite:     txtRefProd.ForeColor = G_ColorBlue
        txtRefAgencia.BackColor = vbWhite:  txtRefAgencia.ForeColor = G_ColorBlue
        txtRefConta.BackColor = vbWhite:    txtRefConta.ForeColor = G_ColorBlue
        fraReferencia.Enabled = True
        lblProdutoRef.Enabled = True
        lblAgenciaRef.Enabled = True
        lblContaRef.Enabled = True
        
        txtReferencia.BackColor = G_ColorGray
        txtReferencia.Enabled = False
        lblRefCliente.Enabled = False
    Else
        txtRefProd.BackColor = G_ColorGray:     txtRefProd.ForeColor = vbBlack
        txtRefAgencia.BackColor = G_ColorGray:  txtRefAgencia.ForeColor = vbBlack
        txtRefConta.BackColor = G_ColorGray:    txtRefConta.ForeColor = vbBlack
        fraReferencia.Enabled = False
        lblProdutoRef.Enabled = False
        lblAgenciaRef.Enabled = False
        lblContaRef.Enabled = False
        
        txtReferencia.BackColor = vbWhite
        txtReferencia.Enabled = True
        lblRefCliente.Enabled = True
    
    End If

End Sub
Private Sub TxtTotal_GotFocus()
    SendKeys "{TAB}"
End Sub
