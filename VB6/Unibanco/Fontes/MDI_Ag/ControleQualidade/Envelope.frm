VERSION 5.00
Begin VB.Form Envelope 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capa de Envelope"
   ClientHeight    =   2784
   ClientLeft      =   168
   ClientTop       =   1644
   ClientWidth     =   9408
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2784
   ScaleWidth      =   9408
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   800
      Left            =   5856
      Picture         =   "Envelope.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   96
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
      Height          =   800
      Left            =   6720
      Picture         =   "Envelope.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
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
      Height          =   800
      Left            =   4128
      Picture         =   "Envelope.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
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
      Height          =   800
      Left            =   4992
      Picture         =   "Envelope.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   800
      Left            =   7584
      Picture         =   "Envelope.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   800
      Left            =   8448
      Picture         =   "Envelope.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   96
      Width           =   850
   End
   Begin VB.Frame fraDadosEnvelope 
      Height          =   1692
      Left            =   96
      TabIndex        =   12
      Top             =   1008
      Width           =   9228
      Begin VB.TextBox txtAgencia 
         BackColor       =   &H8000000F&
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
         Left            =   2784
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   8
         Top             =   432
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
         Left            =   2784
         MaxLength       =   8
         TabIndex        =   0
         Top             =   1056
         Width           =   1428
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
         Left            =   3744
         TabIndex        =   9
         Top             =   480
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
         Left            =   756
         TabIndex        =   7
         Top             =   480
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
         Left            =   720
         TabIndex        =   10
         Top             =   1152
         Width           =   1896
      End
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
      Left            =   720
      TabIndex        =   11
      Top             =   360
      Width           =   1920
   End
   Begin VB.Image imgInformativo 
      Height          =   384
      Left            =   192
      Picture         =   "Envelope.frx":123C
      Top             =   240
      Width           =   384
   End
End
Attribute VB_Name = "Envelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Alterou                  As Boolean
Private mForm                   As Form
Dim sPosicaoErro                As String

Private Type tpModulo
    qryChecarEnvelope           As rdoQuery
    qryChecarAgencia            As rdoQuery
    qryAtualizaCapa             As rdoQuery
    qryInsereAgencia            As rdoQuery
    qryAtualizaDuplicidadeCapa  As rdoQuery
'2109
    qryVerificaCapaRecepcionada As rdoQuery     'Verifica se Capa recepcionada (Status = 0)
    rstModulo                   As rdoResultset
End Type

Private Modulo                  As tpModulo
    

Public Function ShowModal() As Boolean

    ShowModal = False

    Me.Show vbModal

    ShowModal = Alterou

End Function


Private Sub cmdConfirmar_Click()

    On Error GoTo Err_cmdConfirmar
    
    If ValidarDados Then
    
        'Inicia transação para complementar Capa
        Geral.Banco.BeginTrans
    
        'Atualiza dados complementares de capa ou transporta atraves de 'Geral.Capa'
        'dados de capa para ser incluida na tabela Capa
    
        Geral.Capa.AgOrig = txtAgencia.Text
        Geral.Capa.Capa = txtEnvelope.Text
    
        'Trata número de envelope para capa já cadastrada
        If mForm.sCapaOuDocumento = "C" Then
    
            'Atualiza dados da capa e vincula documento com o mesmo Nr Envelope se inserido
'            sPosicaoErro = "AtuCapaEnv"
'            With Modulo.qryAtualizaCapa
'                .rdoParameters(1) = Geral.DataProcessamento
'                .rdoParameters(2) = Geral.Capa.IdCapa
'                .rdoParameters(3) = Geral.Capa.Capa
'                .rdoParameters(4) = Geral.Capa.AgOrig
'                'Vincula Docto
'                .rdoParameters(5) = Geral.Documento.IdDocto
'                'Numero do Malote (0) para Envelope
'                .rdoParameters(6) = 0
'                'Código do CMC7 (Null) para Envelope
'                .rdoParameters(7) = Null
'                .Execute
'
'                If .rdoParameters(0).Value <> 0 Then
'                    MsgBox "Não foi possível atualizar o envelope!", vbExclamation + vbOKOnly, App.Title
'                    txtAgencia.SetFocus: GoTo Exit_cmdConfirmar
'                End If
'            End With
            
            'Verifica se capa com duplicidade
            If Geral.Capa.Duplicidade = 1 Then
                'Atualiza Status e Duplicidade da tabela Capa
                sPosicaoErro = "AtuDuplCapa"
                With Modulo.qryAtualizaDuplicidadeCapa
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Geral.Capa.IdCapa
                    .rdoParameters(3) = Geral.Capa.Status
                    .rdoParameters(4) = Geral.Capa.Duplicidade  'Identificador de envelope em duplicidade
                    .Execute
                    
                    'Verifica se ocorreu erro na atualização
                    If .rdoParameters(0).Value = 1 Then
                        MsgBox "Não foi possível atualizar o envelope!", vbExclamation + vbOKOnly, App.Title
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
    Select Case TratamentoErro(Geral.Banco, "Não foi possível atualizar/inserir o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
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

    'Verifica se Envelope referente a (C)Capa ou (D)Documento em complementação
'    If mForm.sCapaOuDocumento = "C" Then
        'Carrega Agência caso já capturada
        If Val(Geral.AgenciaApresentante) <> 0 Then
            txtAgencia.Text = Format(Geral.AgenciaApresentante, "0000")
            'Call AgenciaValida(txtAgencia, True)
        End If
        
        'Carrega Número do Envelope caso já capturado
        If Geral.Capa.Capa <> 0 Then
            If Geral.Capa.Capa <> 9 Then txtEnvelope.Text = Geral.Capa.Capa
        End If
'    Else
'        'Se Leitura do Documento e menor ou igual a 8, presume-se que é um número de envelope
'        If Len(Trim(Geral.Documento.Leitura)) <= 8 Then
'            txtEnvelope.Text = IIf(Val(Geral.Documento.Leitura) = 0, "", Geral.Documento.Leitura)
'        Else
'            txtEnvelope.Text = ""
'        End If
'    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
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

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = &H1B Then
        CmdSair_Click
    End If
End Sub

Private Sub Form_Load()
    
    txtEnvelope = ""
    Alterou = True
    
    'Limpa Label com descrição da agência
    lblNomeAgencia.Caption = ""
    
    InicializarQuerys
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    With Modulo
        .qryChecarEnvelope.Close
        .qryChecarAgencia.Close
        .qryAtualizaCapa.Close
        .qryInsereAgencia.Close
'2109
        .qryVerificaCapaRecepcionada.Close
        .qryAtualizaDuplicidadeCapa.Close
    End With
    
End Sub

Private Sub txtAgencia_Change()
    
    lblNomeAgencia.Caption = ""

End Sub

Private Sub txtAgencia_GotFocus()
    SelecionarTexto txtAgencia

End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not AgenciaOk Then
            SelecionarTexto txtAgencia
        Else
            SendKeys "{TAB}"
        End If
    End If
    
    SoNumero KeyAscii

End Sub

Private Sub txtAgencia_LostFocus()

'    If Format(Me.ActiveControl.Name, ">") = "CMDSAIR" Then Exit Sub
'    If Trim(txtAgencia.Text) <> "" Then Call AgenciaValida(txtAgencia, True)

End Sub

Private Sub txtEnvelope_GotFocus()
    
    SelecionarTexto txtEnvelope
    
End Sub

Private Sub txtEnvelope_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not EnvelopeOk Then
            SelecionarTexto txtEnvelope
        Else
            'Finaliza digitação
            cmdConfirmar_Click
        End If
    End If
    
    SoNumero KeyAscii

End Sub

Private Sub InicializarQuerys()

    With Modulo
        Set .qryChecarEnvelope = Geral.Banco.CreateQuery("", "{? = CALL MDIAG_ChecarCapaEnvelope (?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Agencia (3)-Nr Capa (4)-Numero de Registros encontrados (5)-IdCapa
            .qryChecarEnvelope.rdoParameters(0).Direction = rdParamReturnValue
            .qryChecarEnvelope.rdoParameters(4).Direction = rdParamOutput
        
        Set .qryAtualizaCapa = Geral.Banco.CreateQuery("", "{? = CALL MDIAG_AtualizaCapa (?,?,?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-IdCapa (3)-Capa (4)-AgOrig (5)-IdDocto (6)-Num_Malote (7)-CMC7
            .qryAtualizaCapa.rdoParameters(0).Direction = rdParamReturnValue
            
        Set .qryInsereAgencia = Geral.Banco.CreateQuery("", "{? = CALL MDIAG_InsereAgencia(?,?,?,?,?,?,?,?)}")
            
        'Atualiza Campo Duplicidade e Status da tabela Capa
        Set .qryAtualizaDuplicidadeCapa = Geral.Banco.CreateQuery("", "{? = CALL MDIAG_AtualizaDuplicidadeCapa (?,?,?,?)}")
            .qryAtualizaDuplicidadeCapa.rdoParameters(0).Direction = rdParamReturnValue
            
        Set .qryChecarAgencia = Geral.Banco.CreateQuery("", "{CALL MDIAG_ObtemAgencia (?)}")
'2109
        'Verifica se Capa recepcionada (Status = 0)
        Set .qryVerificaCapaRecepcionada = Geral.Banco.CreateQuery("", "{? = call MDIAG_VerificaCapaRecepcionada(?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Capa (3)-AgOrig (4)-Num_Malote
            .qryVerificaCapaRecepcionada.rdoParameters(0).Direction = rdParamReturnValue
        
    End With

End Sub

Private Function ValidarDados() As Boolean

ValidarDados = False

On Error GoTo Err_ValidarDados

    'Verifica se agência é válida
    If Not AgenciaOk Then txtAgencia.SetFocus: Exit Function
    
    If Not EnvelopeOk Then txtEnvelope.SetFocus: Exit Function
    
    'Atualiza status de duplicidade de Malote com situação atual
    Geral.Capa.Duplicidade = 0
    
    'Trata número de envelope para capa já cadastrada
    sPosicaoErro = "ChecEnv"
    With Modulo.qryChecarEnvelope
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = txtAgencia.Text
        .rdoParameters(3) = Val(txtEnvelope.Text)
        'Se Capa em Split, verifica IdCapa diferente da atual para evitar o mesmo Nr. de CMC7
        If Geral.Documento.TipoDocto = 1 Then
            .rdoParameters(5) = Geral.Documento.IdCapa
        Else
            .rdoParameters(5) = 0   'Enviar (0) p/ consistir duplicidade independente de IdCapa
        End If
        
        .Execute
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível verificar se Envelope já existe", vbInformation + vbOKOnly, App.Title
            txtEnvelope.SetFocus: Exit Function
        End If

        If .rdoParameters("@Registros") > 0 Then
            If MsgBox("O Número do Envelope já existe, deseja recadastrá-lo ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Geral.Capa.Duplicidade = 1
            Else
               txtEnvelope.SetFocus: Exit Function
            End If
        End If
    End With

    ValidarDados = True
                
Exit Function
                
'    '**************************************
'    '* Chama a função que busca a agência *
'    '* do envelope, para tratar a tabela  *
'    '* de feriados. Leda (09/05/2000)     *
'    '**************************************
'    BuscaAgenciaEnvelope
'
'    Exit Function
    
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
  
    Select Case TratamentoErro(Geral.Banco, "Não foi possível validar o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    Me.Hide

End Function
Private Function EnvelopeOk() As Boolean
     
    EnvelopeOk = False

     Dim sEnvelope As String
    
     sEnvelope = CStr(Val(txtEnvelope))
     
     'Verifica se o Número do envelope é válido
     If Len(Trim(txtEnvelope.Text)) < 2 Then
        Beep
        MsgBox "O Número do envelope deve ter pelo menos 2 dígitos!", vbExclamation + vbOKOnly, App.Title
        GoTo Sair
     ElseIf Right(sEnvelope, 1) <> Modulo11UBB(Val(Left(sEnvelope, Len(sEnvelope) - 1))) Then
         If Right(sEnvelope, 1) <> Modulo11Simplificado(Val(Left(sEnvelope, Len(sEnvelope) - 1))) Then
             If Right(sEnvelope, 1) <> Modulo11U(Val(Left(sEnvelope, Len(sEnvelope) - 1))) Then
                Beep
                MsgBox "Dígito verificador não confere", vbExclamation + vbOKOnly, App.Title
                GoTo Sair
             End If
         End If
     End If
    
    EnvelopeOk = True

Sair:

End Function
Private Function AgenciaOk() As Boolean
    Dim iErroData As Integer
    
    AgenciaOk = False
    
    'Verifica se agência é válida
    If Len(Trim(txtAgencia)) = 0 Then
        MsgBox "A Agência de origem deve ser informada!", vbExclamation + vbOKOnly, App.Title
        GoTo Sair
    End If
    If Not AgenciaValida(txtAgencia.Text) Then
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
Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

