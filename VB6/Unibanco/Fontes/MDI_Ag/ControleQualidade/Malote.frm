VERSION 5.00
Begin VB.Form Malote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capa de Malote"
   ClientHeight    =   3180
   ClientLeft      =   384
   ClientTop       =   1320
   ClientWidth     =   10632
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   10632
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
      Height          =   800
      Left            =   7920
      Picture         =   "Malote.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   96
      Width           =   852
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
      Height          =   800
      Left            =   7056
      Picture         =   "Malote.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   96
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
      Height          =   800
      Left            =   6192
      Picture         =   "Malote.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdCMC7 
      Caption         =   "C&MC7"
      Height          =   800
      Left            =   3600
      Picture         =   "Malote.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   800
      Left            =   4464
      Picture         =   "Malote.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   800
      Left            =   5328
      Picture         =   "Malote.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   800
      Left            =   9648
      Picture         =   "Malote.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   800
      Left            =   8784
      Picture         =   "Malote.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   96
      Width           =   850
   End
   Begin VB.Frame fraDadosEnvelope 
      Height          =   2124
      Left            =   96
      TabIndex        =   20
      Top             =   1008
      Width           =   10380
      Begin VB.Frame fraCMC7 
         Caption         =   "Linha do CMC7"
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
         Height          =   780
         Left            =   5568
         TabIndex        =   2
         Top             =   144
         Width           =   4284
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
            Height          =   408
            Left            =   192
            MaxLength       =   8
            TabIndex        =   3
            Top             =   240
            Width           =   1044
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
            Height          =   408
            Left            =   2640
            MaxLength       =   12
            TabIndex        =   5
            Top             =   240
            Width           =   1452
         End
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
            Height          =   408
            Left            =   1308
            MaxLength       =   10
            TabIndex        =   4
            Top             =   240
            Width           =   1236
         End
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
         Left            =   3456
         MaxLength       =   14
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   384
         Width           =   1620
      End
      Begin VB.TextBox txtContaMalote 
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
         Left            =   3456
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1536
         Width           =   1968
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
         Left            =   3456
         MaxLength       =   4
         TabIndex        =   7
         Top             =   960
         Width           =   756
      End
      Begin VB.Label lblNumeroMalote 
         AutoSize        =   -1  'True
         Caption         =   "Número do Malote Empresa:"
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
         Left            =   816
         TabIndex        =   9
         Top             =   1584
         Width           =   2532
      End
      Begin VB.Label lblNumeroEnvelope 
         AutoSize        =   -1  'True
         Caption         =   "Número Capa Malote:"
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
         Left            =   768
         TabIndex        =   0
         Top             =   432
         Width           =   1932
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
         Left            =   804
         TabIndex        =   6
         Top             =   1008
         Width           =   1776
      End
      Begin VB.Label lblNomeAgencia 
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
         Left            =   4368
         TabIndex        =   8
         Top             =   1008
         Width           =   5148
      End
   End
   Begin VB.Label lblInformativo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de Malote Empresa"
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
      Left            =   720
      TabIndex        =   19
      Top             =   408
      Width           =   2484
   End
   Begin VB.Image imgInformativo 
      Height          =   384
      Left            =   192
      Picture         =   "Malote.frx":1850
      Top             =   288
      Width           =   384
   End
End
Attribute VB_Name = "Malote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Alterou As Boolean
Private mForm As Form
Dim sPosicaoErro As String

Private Type tpModulo
    qryChecarEnvelope As rdoQuery
    qryAtualizaCapa As rdoQuery
    qryChecarAgencia As rdoQuery
    qryAtualizaDuplicidadeCapa As rdoQuery
    qryInsereAgencia As rdoQuery
'2109
    qryVerificaCapaRecepcionada As rdoQuery     'Verifica se Capa recepcionada (Status = 0)
    
    rstModulo As rdoResultset
    DuplicidadeCapa As Integer    'Situação atual da capa de Malote na Tabela CAPA
End Type

Private Modulo As tpModulo

Private Sub cmdCMC7_Click()
    
    Call HabilitaCMC7(True)
    
    cmdCMC7.Enabled = False
    txtCMC71.SetFocus
    'Para digitação do CMC7 limpa campo contendo Número do Malote e guarda situação anterior
    txtEnvelope.Tag = txtEnvelope.Text
    txtEnvelope.Text = ""
    txtCMC71.Tag = txtCMC71.Text
    txtCMC72.Tag = txtCMC72.Text
    txtCMC73.Tag = txtCMC73.Text
    
End Sub

Private Sub cmdConfirmar_Click()

Dim bAtualizouCMC7 As Boolean

    bAtualizouCMC7 = False
    
    On Error GoTo Err_cmdConfirmar
    
    If ValidarDados Then
        
        'Inicia transação para complementar Capa
        Geral.Banco.BeginTrans
        
        'Atualiza dados complementares de capa ou transporta atraves de 'Geral.Capa'
        'dados de capa para ser incluida na tabela Capa
        
        Geral.Capa.AgOrig = txtAgencia.Text
        Geral.Capa.Capa = Val(txtEnvelope.Text)
        Geral.Capa.Num_Malote = txtContaMalote.Text
        If (txtCMC71.Text + txtCMC72.Text + txtCMC73.Text) <> "" Then
                Geral.Documento.Leitura = Format(txtEnvelope.Text, "00000000000000")
                bAtualizouCMC7 = True
        End If
        
        'Trata número de envelope para capa já cadastrada
        If mForm.sCapaOuDocumento = "C" Then
        
            'Atualiza dados da capa e vincula documento com o mesmo Nr Envelope se inserido
            sPosicaoErro = "AtuCapaEnv"
            With Modulo.qryAtualizaCapa
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Geral.Capa.IdCapa
                .rdoParameters(3) = Geral.Capa.Capa
                .rdoParameters(4) = Geral.Capa.AgOrig
                'Vincula Docto
                .rdoParameters(5) = Geral.Documento.IdDocto
                'Numero do Malote
                .rdoParameters(6) = Geral.Capa.Num_Malote
                'Código do CMC7
                .rdoParameters(7) = Geral.Documento.Leitura
                .Execute
                
                If .rdoParameters(0).Value <> 0 Then
                    Beep
                    MsgBox "Não foi possível atualizar o malote!", vbExclamation + vbOKOnly, App.Title
                    txtAgencia.SetFocus: GoTo Exit_cmdConfirmar
                End If
            End With
            
            'Atualiza flag de Duplicidade da Capa
            If Geral.Capa.Duplicidade = 1 Then
                'Atualiza Status e Duplicidade da tabela Capa
                sPosicaoErro = "AtuDuplCapa"
                With Modulo.qryAtualizaDuplicidadeCapa
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Geral.Capa.IdCapa
                    .rdoParameters(3) = Geral.Capa.Status
                    .rdoParameters(4) = Geral.Capa.Duplicidade  'Identificador de malote em duplicidade
                    .Execute
                    
                    'Verifica se ocorreu erro na atualização
                    If .rdoParameters(0).Value = 1 Then
                        MsgBox "Não foi possível atualizar o malote!", vbExclamation + vbOKOnly, App.Title
                        txtAgencia.SetFocus: GoTo Exit_cmdConfirmar
                    End If
                End With
            End If
        End If
        
    '2109
        'Verifica se existe Capa Recepcionada
        sPosicaoErro = "VerifCapaRecep"
        With Modulo.qryVerificaCapaRecepcionada
            .rdoParameters(1) = Geral.DataProcessamento
            .rdoParameters(2) = Geral.Capa.Capa
            .rdoParameters(3) = Geral.Capa.AgOrig
            .rdoParameters(4) = Geral.Capa.Num_Malote
            .Execute
            
            'Verifica se não existe capa apenas recepcionada
            If .rdoParameters(0).Value = 0 Then
        
                'Grava informações da agencia do envelope
                sPosicaoErro = "InsAgencia"
                With Modulo.qryInsereAgencia
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Val(txtAgencia)
                    .rdoParameters(3) = 0           'lacre
                    .rdoParameters(4) = 1           'qtd informada
                    .rdoParameters(5) = 1           'qtd gravada
                    .rdoParameters(6) = "complem."  'identificador
                    .rdoParameters(7) = Format(Now, "hh:mm")    'hora chegada
                    .rdoParameters(8) = Format(Now, "hh:mm")    'hora cadastrada
                    .rdoParameters(9) = Geral.Capa.IdEnv_Mal  'Indicador (E)Envelope / (M)Malote
                    .Execute
                    'Verifica se ocorreu erro na atualização
                    If .rdoParameters(0).Value <> 0 Then
                        Beep
                        MsgBox "Não foi possível atualizar agência!", vbExclamation + vbOKOnly, App.Title
                        txtAgencia.SetFocus: GoTo Exit_cmdConfirmar
                    End If
                End With
            End If
        End With
        
        'Finaliza Transação
        Geral.Banco.CommitTrans
        
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
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
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

    'Carrega campo Duplicidade da tabela CAPA (Utilizado na verificação de Duplicidade de Capa)
    If mForm.sCapaOuDocumento = "C" Then
        Modulo.DuplicidadeCapa = Geral.Capa.Duplicidade
    Else
        Modulo.DuplicidadeCapa = 0  'Para todo Documento com número de Capa em duplicidade na Tabela CAPA, assume-se capa em Duplicidade
    End If

    'Formata o Número do CMC7 OU Capa de Malote
    If Len(Trim(Geral.Capa.Capa)) <> 13 Or _
        Left(Format(Geral.Capa.Capa, "00000000000000"), 4) <> "0600" Then
        If Trim(Geral.Documento.Leitura) = "9" Then Geral.Documento.Leitura = ""
        txtCMC71.Text = IIf(Val(Geral.Documento.Leitura) = 0, "", Left(Geral.Documento.Leitura, 8))
        txtCMC72.Text = IIf(Val(Mid(Geral.Documento.Leitura, 9, 10)) = 0, "", Mid(Geral.Documento.Leitura, 9, 10))
        txtCMC73.Text = IIf(Val(Mid(Geral.Documento.Leitura, 19)) = 0, "", Mid(Geral.Documento.Leitura, 19))
        
        If Len(Geral.Documento.Leitura) = 30 And _
            Mid(Geral.Documento.Leitura, 1, 3) = "409" And _
            Mid(Geral.Documento.Leitura, 19, 4) = "0600" And _
            Mid(Geral.Documento.Leitura, 18, 1) = "4" And _
            Mid(Geral.Documento.Leitura, 9, 3) = "600" And _
            Mid(Geral.Documento.Leitura, 12, 6) = Mid(Geral.Documento.Leitura, 24, 6) Then
            'Formata o Número da Capa de Malote
            txtEnvelope.Text = Mid(Geral.Documento.Leitura, 19, 4) & Mid(Geral.Documento.Leitura, 12, 6) & Mid(Geral.Documento.Leitura, 4, 4)
        Else
            Call HabilitaCMC7(True)
        End If
    Else
        txtEnvelope.Text = Format(Geral.Capa.Capa, "00000000000000")
    End If
    
    'Verifica se Envelope referente a (C)Capa ou (D)Documento em complementação
    If mForm.sCapaOuDocumento = "C" Then
        
        'Carrega Agência caso já capturada
        If Geral.Capa.AgOrig <> 0 Then
            txtAgencia.Text = Format(Geral.Capa.AgOrig, "0000")
            Call AgenciaValida(txtAgencia, True)
        End If
        
        'Se Numero de Malote já existe então desabilita, senão habilita para digitação
        If Geral.Capa.Num_Malote <> 0 Then
            txtContaMalote.Text = Geral.Capa.Num_Malote
        End If
        
    End If
    
    'Guarda situação anterior do CMC7 e Capa de malote
    txtEnvelope.Tag = txtEnvelope.Text
    txtCMC71.Tag = txtCMC71.Text
    txtCMC72.Tag = txtCMC72.Text
    txtCMC73.Tag = txtCMC73.Text
    If fraCMC7.Enabled Then
'        txtCMC71.BackColor = vbWhite: txtCMC72.BackColor = vbWhite: txtCMC73.BackColor = vbWhite
        cmdCMC7.Enabled = False
        txtCMC71.SelStart = 0: txtCMC71.SelLength = txtCMC71.MaxLength
        txtCMC71.SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = &H1B Then
        If InStr(Format(ActiveControl.Name, ">"), "TXTCMC7") <> 0 Then Exit Sub
        CmdSair_Click
    End If
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
    
    txtEnvelope = ""
    txtEnvelope.ForeColor = vbBlack
    txtEnvelope.BackColor = G_ColorGray
    txtEnvelope.Locked = True
    
    Alterou = True
    
    'Limpa Label com descrição da agência
    lblNomeAgencia.Caption = ""

    Call HabilitaCMC7(False)
    cmdCMC7.Enabled = True
    
    'Inicializar todas query's
    InicializarQuery
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    With Modulo
        .qryChecarEnvelope.Close
        .qryAtualizaCapa.Close
        .qryChecarAgencia.Close
        .qryAtualizaDuplicidadeCapa.Close
        .qryInsereAgencia.Close
'2109
        .qryVerificaCapaRecepcionada.Close
        
    End With
    
End Sub

Private Sub txtAgencia_Change()
    
    lblNomeAgencia.Caption = ""
    
End Sub

Private Sub txtAgencia_GotFocus()

txtAgencia.SelStart = 0
txtAgencia.SelLength = txtAgencia.MaxLength

End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

InibirTeclaAlfa KeyAscii

If KeyAscii = 13 Then
    KeyAscii = 0
    
    If Not AgenciaOk Then
        txtAgencia.SelStart = 0
        txtAgencia.SelLength = txtAgencia.MaxLength
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub txtAgencia_LostFocus()

    If Format(Me.ActiveControl.Name, ">") = "CMDSAIR" Then Exit Sub
    If Trim(txtAgencia.Text) <> "" Then Call AgenciaValida(txtAgencia, True)

End Sub

Private Sub txtCMC71_Change()
    
    If txtCMC71.Enabled Then
        If Len(Trim(txtCMC71.Text)) = txtCMC71.MaxLength Then SendKeys "{ENTER}"
    End If
    
End Sub

Private Sub txtCMC71_GotFocus()

    txtCMC71.SelStart = 0
    txtCMC71.SelLength = txtCMC71.MaxLength

End Sub

Private Sub txtCMC71_KeyPress(KeyAscii As Integer)

    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC71) < 8 Then
            Beep
            MsgBox "Digite todos os números do primeiro campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            txtCMC71.SelStart = 0
            txtCMC71.SelLength = txtCMC71.MaxLength
            GoTo Sair
        End If
        'Verifica o código do documento
        If Left(txtCMC71.Text, 3) <> "409" Then
            Beep
            MsgBox "Para Capa de Malote o primeiro campo do CMC7 deve começar com o Nr. 409.", vbInformation + vbOKOnly, App.Title
            GoTo Sair
        End If
        txtCMC72.SetFocus
        
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        'Desabilita digitação de CMC7
        Call HabilitaCMC7(False)
        cmdCMC7.Enabled = True
        'Se cancelada digitação do CMC7, retorna situação anterior
        txtEnvelope.Text = txtEnvelope.Tag
        txtCMC71.Text = txtCMC71.Tag
        txtCMC72.Text = txtCMC72.Tag
        txtCMC73.Text = txtCMC73.Tag
    End If
    
    Exit Sub
    
Sair:
    txtCMC71.SelStart = 0
    txtCMC71.SelLength = txtCMC71.MaxLength
    txtCMC71.SetFocus

End Sub

Private Sub txtCMC72_Change()
    
    If txtCMC72.Enabled Then
        If Len(Trim(txtCMC72.Text)) = txtCMC72.MaxLength Then SendKeys "{ENTER}"
    End If
    
End Sub

Private Sub txtCMC72_GotFocus()
    
    txtCMC72.SelStart = 0
    txtCMC72.SelLength = txtCMC72.MaxLength

End Sub

Private Sub txtCMC72_KeyPress(KeyAscii As Integer)

    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC72) < 10 Then
            Beep
            MsgBox "Digite todos os números do segundo campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            GoTo Sair
        End If
        'Verifica o código do documento
        If Left(txtCMC72.Text, 3) <> "600" Then
            Beep
            MsgBox "Para Capa de Malote o segundo campo do CMC7 deve começar com o Nr. 600.", vbInformation + vbOKOnly, App.Title
            GoTo Sair
        End If
        txtCMC73.SetFocus
        
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        'Desabilita digitação de CMC7
        Call HabilitaCMC7(False)
        cmdCMC7.Enabled = True
        'Se cancelada digitação do CMC7, retorna situação anterior
        txtEnvelope.Text = txtEnvelope.Tag
        txtCMC71.Text = txtCMC71.Tag
        txtCMC72.Text = txtCMC72.Tag
        txtCMC73.Text = txtCMC73.Tag
    End If

    Exit Sub
    
Sair:
    txtCMC72.SelStart = 0
    txtCMC72.SelLength = txtCMC72.MaxLength
    txtCMC72.SetFocus

End Sub

Private Sub txtCMC73_Change()

    If txtCMC73.Enabled Then
        If Len(Trim(txtCMC73.Text)) = txtCMC73.MaxLength Then SendKeys "{ENTER}"
    End If
    
End Sub

Private Sub txtCMC73_GotFocus()
    
    txtCMC73.SelStart = 0
    txtCMC73.SelLength = txtCMC73.MaxLength

End Sub

Private Sub txtCMC73_KeyPress(KeyAscii As Integer)
    
    Dim iErroCMC7 As Integer
    
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC73) < 12 Then
            Beep
            MsgBox "Digite todos os números do terceiro campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            GoTo SairCMC73
        End If
        
        If Left(txtCMC73, 4) <> "0600" Then
            Beep
            MsgBox "Para Capa de Malote o terceiro campo do CMC7 deve começar com o Nr. 0600.", vbInformation + vbOKOnly, App.Title
            GoTo SairCMC73
        End If
        
        If Not CMC7Ok(iErroCMC7) Then
            Beep
            MsgBox Switch(iErroCMC7 = 1, "Primeiro", iErroCMC7 = 2, "Segundo", iErroCMC7 = 3, "Terceiro") & _
                    " campo do CMC7 não confere!", vbExclamation + vbOKOnly, App.Title
            If iErroCMC7 = 1 Then
                txtCMC71.SetFocus: GoTo Sair
            End If
            If iErroCMC7 = 2 Then
                txtCMC72.SetFocus: GoTo Sair
            End If
            If iErroCMC7 = 3 Then
                GoTo SairCMC73
            End If
        Else
            'Forma o número da Capa de Malote
            txtEnvelope.Text = Left(txtCMC73.Text, 4) + Mid(txtCMC72.Text, 4, 6) + Mid(txtCMC71.Text, 4, 4)
            'Desabilita digitação do CMC7
            Call HabilitaCMC7(False)
            cmdCMC7.Enabled = True
        End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        'Desabilita digitação de CMC7
        Call HabilitaCMC7(False)
        cmdCMC7.Enabled = True
        'Se cancelada digitação do CMC7, retorna situação anterior
        txtEnvelope.Text = txtEnvelope.Tag
        txtCMC71.Text = txtCMC71.Tag
        txtCMC72.Text = txtCMC72.Tag
        txtCMC73.Text = txtCMC73.Tag
    End If

    Exit Sub

SairCMC73:
    txtCMC73.SelStart = 0
    txtCMC73.SelLength = txtCMC73.MaxLength
    txtCMC73.SetFocus
    
Sair:

End Sub

Private Sub txtContaMalote_GotFocus()

txtContaMalote.SelStart = 0
txtContaMalote.SelLength = txtContaMalote.MaxLength

End Sub

Private Sub txtContaMalote_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfa KeyAscii

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not ContaMaloteOk Then
            txtContaMalote.SelStart = 0
            txtContaMalote.SelLength = txtContaMalote.MaxLength
        Else
            'Finaliza digitação
            cmdConfirmar_Click
        End If
    End If
    
End Sub
Private Function ValidarDados()

ValidarDados = False

On Error GoTo Err_ValidarDados

    'Verifica se Capa de Malote é válido
    If Not EnvelopeOk Then
        If cmdCMC7.Enabled Then
            cmdCMC7_Click
        Else
            txtCMC71.SetFocus
        End If
        Exit Function
    End If
    
    'Verifica se Agência é válida
    If Not AgenciaOk Then txtAgencia.SetFocus: Exit Function
    
    'Atualiza status de duplicidade de Malote com situação atual
    Geral.Capa.Duplicidade = Modulo.DuplicidadeCapa
    
    'Trata número de Malote para capa já cadastrada
    sPosicaoErro = "ChecMal"
    With Modulo.qryChecarEnvelope
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Null    ' Para malote não há necessidade de verificar por agência
        .rdoParameters(3) = Val(txtEnvelope.Text)
        'Se Capa em Split, verifica IdCapa diferente da atual para evitar o mesmo Nr. de CMC7
        If Geral.Documento.TipoDocto = 1 Then
            .rdoParameters(5) = Geral.Documento.IdCapa
        Else
            .rdoParameters(5) = 0   'Enviar (0) p/ consistir duplicidade independente de IdCapa
        End If
        .Execute
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível verificar se Capa de Malote já existe", vbInformation + vbOKOnly, App.Title
            txtAgencia.SetFocus: Exit Function
        End If

        'Se existe Capa com duplicidade, solicita recadastramento e envia para supervisor
        If mForm.sCapaOuDocumento = "D" And .rdoParameters("@Registros") > 0 Then
            If MsgBox("Capa de Malote já existe, deseja recadastrá-la ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Geral.Capa.Duplicidade = 1
            Else
               txtAgencia.SetFocus: Exit Function
            End If
        Else
            
            If .rdoParameters("@Registros") = 0 Then
                Geral.Capa.Duplicidade = 0
            Else
                Geral.Capa.Duplicidade = 1
            End If
        
'            'Verifica se retornou apenas um registro com o mesmo número de Capa
'            If .rdoParameters("@Registros") = 1 Then
'                'Verifica se duplicidade trata-se da mesma capa de Malote
'                If (txtEnvelope.Tag = txtEnvelope.Text) Then
'                    Geral.capa.Duplicidade = 0
'                Else
'                    Geral.capa.Duplicidade = 1
'                End If
'            End If
'            'Verifica se retornou Mais de um registro com o mesmo número de Capa
'            If .rdoParameters("@Registros") > 1 Then
'                'Verifica se duplicidade trata-se da mesma capa de Malote, permanece
'                'a mesma situação de duplicidade da capa
'                If (txtEnvelope.Tag = txtEnvelope.Text) Then
'                    Geral.capa.Duplicidade = Modulo.DuplicidadeCapa
'                Else
'                    Geral.capa.Duplicidade = 1
'                End If
'            End If
            
            If (txtEnvelope.Tag <> txtEnvelope.Text) And Geral.Capa.Duplicidade = 1 Then
                If MsgBox("Capa de Malote já existe, deseja recadastrá-la ?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                    'Se não autorizou recadastramento de capa, retorna situação anterior
                    Geral.Capa.Duplicidade = Modulo.DuplicidadeCapa
                    txtAgencia.SetFocus: Exit Function
                End If
            End If

        End If
    End With

    'Verifica se agência é válida
    If Not AgenciaOk Then txtAgencia.SetFocus: Exit Function
    
    'Verifica se Número do Malote Empresa é válido
    If Not ContaMaloteOk Then txtContaMalote.SetFocus: Exit Function
    
    ValidarDados = True
                
Exit Function

Err_ValidarDados:
   
    Alterou = False
  
    Select Case TratamentoErro("Não foi possível validar o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    Me.Hide

End Function
Private Function AgenciaValida(ByVal iCdAgencia As Integer, Optional bMostraAgencia As Boolean = False) As Boolean

On Error GoTo Err_AgenciaValida

AgenciaValida = False

sPosicaoErro = "ChecAgencia"
With Modulo.qryChecarAgencia
    .rdoParameters(0) = iCdAgencia
    Set Modulo.rstModulo = .OpenResultset(rdOpenStatic)
End With

If Modulo.rstModulo.RowCount > 0 Then
    AgenciaValida = True
Else
    GoTo Exit_AgenciaValida
End If

If bMostraAgencia Then lblNomeAgencia.Caption = Modulo.rstModulo!agefsnoagen

Exit_AgenciaValida:

    If Not (Modulo.rstModulo Is Nothing) Then Modulo.rstModulo.Close
    Exit Function
    
Err_AgenciaValida:
   
    Alterou = False
    
    'Fecha Resultset
    If Not (Modulo.rstModulo Is Nothing) Then Modulo.rstModulo.Close
  
    Select Case TratamentoErro("Não foi possível validar o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    Me.Hide

End Function

Private Function AgenciaOk() As Boolean
    Dim iErroData  As Integer
    
    AgenciaOk = False
        
    'Verifica se agência é válida
    If Len(Trim(txtAgencia)) = 0 Then
        Beep
        MsgBox "A Agência de origem deve ser informada!", vbExclamation + vbOKOnly, App.Title
        GoTo Sair
    End If
    If Not AgenciaValida(txtAgencia.Text) Then
        Beep
        MsgBox "Código de Agência inválido. Verifique!", vbExclamation, "Agência"
        GoTo Sair
    End If
    
    'Verifica se feriado na agência
    iErroData = ValidaAgencia(CInt(txtAgencia.Text), "", False)
    If iErroData <> 0 Then
        Select Case iErroData
            Case 2 'Feriado
                MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
            Case 3 'Agência Fechada
                MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
        End Select
        If iErroData = 2 Or iErroData = 3 Then
            GoTo Sair
        End If
    End If
    
    AgenciaOk = True

Sair:

End Function

Private Function EnvelopeOk() As Boolean
     
    EnvelopeOk = False

     Dim sEnvelope As String
    
     sEnvelope = CStr(Val(txtEnvelope))
     
     'Verifica se o Número do envelope é válido
    If Val(txtEnvelope) = 0 Then
        Beep
        MsgBox "Entre com o CMC7 para obter o Número da Capa de Malote!", vbExclamation, App.Title
        GoTo Sair
    ElseIf Len(txtEnvelope) <> 14 Then
        Beep
        MsgBox "O Número da Capa do Malote deverá conter 14 digitos .", vbExclamation, App.Title
        GoTo Sair
    ElseIf Mid(Format(txtEnvelope, "00000000000000"), 1, 4) <> "0600" Then
        Beep
        MsgBox "O Número da Capa do Malote deverá iniciar com 0600 !", vbExclamation, App.Title
        GoTo Sair
    End If

    EnvelopeOk = True

Sair:

End Function

Private Function ContaMaloteOk() As Boolean
    Dim iLenMalote As Integer
    
    ContaMaloteOk = False
    iLenMalote = Len(Trim(txtContaMalote))
    
    If iLenMalote = 0 Then
        Beep
        MsgBox "Digite o Número do Malote Empresa !", vbExclamation + vbOKOnly, App.Title
        GoTo Sair
    End If
   
    'Verifica Nr.Malote Novo ou Antigo
    If (Left(CStr(txtContaMalote), 2) = "09" And iLenMalote = 12) Or _
        (Left(CStr(txtContaMalote), 1) = "9" And iLenMalote = 11) Then
        txtContaMalote = Format(txtContaMalote, "000000000000")
        If Val(Mid(txtContaMalote, 3, 9)) < 1 Then
            GoTo SairComMsg
        End If
        If Left(CStr(txtContaMalote), 2) <> "09" Then GoTo SairComMsg
    Else
        If iLenMalote > 11 Then GoTo SairComMsg
    
        txtContaMalote = Format(txtContaMalote, "00000000000")
    End If

    'Calcula Modulo 10 para Nr Malote antigo (11) ou Novo (12)posições
    If Not Modulo10(txtContaMalote, Len(txtContaMalote.Text)) Then
        GoTo SairComMsg
    End If
    
    ContaMaloteOk = True

Sair:
    Exit Function
    
SairComMsg:
    Beep
    MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title

End Function
Private Sub InicializarQuery()
    
    With Modulo
        Set .qryChecarEnvelope = Geral.Banco.CreateQuery("", "{? = call ChecarCapaEnvelope  (?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Agencia (3)-Nr Capa (4)-Numero de Registros encontrados (5)-IdCapa
            .qryChecarEnvelope.rdoParameters(0).Direction = rdParamReturnValue
            .qryChecarEnvelope.rdoParameters(4).Direction = rdParamOutput
        
        Set .qryAtualizaCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaCapa (?,?,?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-IdCapa (3)-Capa (4)-AgOrig (5)-IdDocto (6)-Num_Malote (7)-CMC7
            .qryAtualizaCapa.rdoParameters(0).Direction = rdParamReturnValue
        
        'Atualiza Campo Duplicidade e Status da tabela Capa
        Set .qryAtualizaDuplicidadeCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaDuplicidadeCapa (?,?,?,?)}")
            .qryAtualizaDuplicidadeCapa.rdoParameters(0).Direction = rdParamReturnValue
        
        Set .qryInsereAgencia = Geral.Banco.CreateQuery("", "{? = call InsereAgencia(?,?,?,?,?,?,?,?,?)}")

        Set .qryChecarAgencia = Geral.Banco.CreateQuery("", "{call ObtemAgencia (?)}")
'2109
        'Verifica se Capa recepcionada (Status = 0)
        Set .qryVerificaCapaRecepcionada = Geral.Banco.CreateQuery("", "{? = call VerificaCapaRecepcionada(?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Capa (3)-AgOrig (4)-Num_Malote
            .qryVerificaCapaRecepcionada.rdoParameters(0).Direction = rdParamReturnValue

    End With

End Sub
Private Function CMC7Ok(ByRef iCampoCMC7Erro As Integer) As Boolean
'-------------------------------------------------------------------------------------
'       Validar campos de digitação do CMC7
'
' Parâmetros:   iCampoCMC7Erro    - Número do campo CMC7 com erro de verificação
'
' Retorno:      True    - CMC7 ok
'               False   - CMC7 com erro de dígito
'-------------------------------------------------------------------------------------
    
    Dim sCMC7 As String, sCmc71 As String, sCmc72 As String
    Dim sCmc73 As String, svalor As String
    
    CMC7Ok = False: iCampoCMC7Erro = 0
    
    'Verifica o código do documento
    If Left(txtCMC72.Text, 3) <> "600" Then iCampoCMC7Erro = 1: GoTo Sair
    If Left(txtCMC73, 4) <> "0600" Then iCampoCMC7Erro = 3: GoTo Sair
    
    sCMC7 = Format(txtCMC71, "00000000") + Format(txtCMC72, "0000000000") + Format(txtCMC73, "000000000000")
    If Not TratarCamposCMC7(sCMC7, sCmc71, sCmc72, sCmc73, svalor) Then
        If Val(sCmc71) = 0 Then iCampoCMC7Erro = 1: GoTo Sair
        If Val(sCmc72) = 0 Then iCampoCMC7Erro = 2: GoTo Sair
        If Val(sCmc73) = 0 Then iCampoCMC7Erro = 3: GoTo Sair
    End If
    
    CMC7Ok = True

Sair:

End Function
Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub


Private Sub HabilitaCMC7(bHabilita As Boolean)

    If bHabilita Then
        fraCMC7.Enabled = True
        txtCMC71.BackColor = vbWhite: txtCMC71.ForeColor = G_ColorBlue
        txtCMC72.BackColor = vbWhite: txtCMC72.ForeColor = G_ColorBlue
        txtCMC73.BackColor = vbWhite: txtCMC73.ForeColor = G_ColorBlue
    Else
        fraCMC7.Enabled = False
        txtCMC71.BackColor = G_ColorGray: txtCMC71.ForeColor = vbBlack
        txtCMC72.BackColor = G_ColorGray: txtCMC72.ForeColor = vbBlack
        txtCMC73.BackColor = G_ColorGray: txtCMC73.ForeColor = vbBlack
    End If
    
End Sub
