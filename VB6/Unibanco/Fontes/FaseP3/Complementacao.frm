VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Complementacao 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complementação"
   ClientHeight    =   8976
   ClientLeft      =   60
   ClientTop       =   336
   ClientWidth     =   12192
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8976
   ScaleWidth      =   12192
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picAguardo 
      Height          =   2556
      Left            =   2640
      ScaleHeight     =   2508
      ScaleWidth      =   7020
      TabIndex        =   8
      Top             =   2460
      Width           =   7068
      Begin VB.CommandButton cmdFechar 
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   2688
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1596
      End
      Begin ComctlLib.ProgressBar pgbAguardo 
         Height          =   348
         Left            =   576
         TabIndex        =   10
         Top             =   960
         Width           =   5868
         _ExtentX        =   10351
         _ExtentY        =   614
         _Version        =   327682
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.Label lblAguardo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aguardando documento para complementação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1152
         TabIndex        =   11
         Top             =   672
         Width           =   4848
      End
   End
   Begin VB.Timer TmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   288
      Top             =   6096
   End
   Begin MSFlexGridLib.MSFlexGrid grdDocumentos 
      Height          =   588
      Left            =   288
      TabIndex        =   7
      Top             =   5136
      Visible         =   0   'False
      Width           =   3804
      _ExtentX        =   6710
      _ExtentY        =   1037
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.PictureBox picLabelData 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   9216
      ScaleHeight     =   252
      ScaleWidth      =   2892
      TabIndex        =   5
      Top             =   48
      Width           =   2940
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2904
      End
   End
   Begin VB.PictureBox picLabeCapa 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   3504
      ScaleHeight     =   252
      ScaleWidth      =   5628
      TabIndex        =   3
      Top             =   48
      Width           =   5676
      Begin VB.Label lblCapa 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -48
         TabIndex        =   4
         Top             =   0
         Width           =   5676
      End
   End
   Begin VB.PictureBox picLabelDocumento 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   48
      ScaleHeight     =   252
      ScaleWidth      =   3372
      TabIndex        =   1
      Top             =   48
      Width           =   3420
      Begin VB.Label lblDocumento 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3384
      End
   End
   Begin LeadLib.Lead Lead1 
      Height          =   4512
      Left            =   96
      TabIndex        =   0
      Top             =   384
      Visible         =   0   'False
      Width           =   12012
      _Version        =   524288
      _ExtentX        =   21188
      _ExtentY        =   7959
      _StockProps     =   229
      BackColor       =   16777215
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
      ScaleHeight     =   374
      ScaleWidth      =   999
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
   End
End
Attribute VB_Name = "Complementacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tpModulo    'Resultset de Documentos
    qryDadosCapa                        As rdoQuery
    tbDocumentos                        As rdoResultset       'Capa a digitar
    qryGetIdCapaComplementacao          As rdoQuery           'Retorna capa para complementação
    qryGetTodosDocumentosCapa           As rdoQuery           'Obtem todos documentos referentes a uma Capa
    qryChecarEnvelope                   As rdoQuery           'Retorna quantos números de capa existem com a mesma agencia na tabela capa
    qryAtualizaDuplicidadeCapa          As rdoQuery           'Atualiza Campo Duplicidade e Status da tabela Capa
    qryInsereNovaCapa                   As rdoQuery           'Insere nova capa de envelope/Malote
    qryAtualizaArrecEletronica          As rdoQuery           'Insere Dados de Arrecadação
    qryGetChequeDuplicado               As rdoQuery           'Verifica duplicidade da numeração de cheque
    qryAtualizaDocumentoExcluido        As rdoQuery           'Move documento cheque para duplicado
    qryAtualizaCheque                   As rdoQuery           'Insere dados de cheque na tabela cheque
    qryRemoveCapaRecepcionada           As rdoQuery           'Remove Capa apenas recepcionada (Status = 0)
    qryGetIdDocto                       As rdoQuery           'Obtem dados de Documento da Capa anterior
    qrySplitCapaAnterior                As rdoQuery           'Faz o Split de Capa Atual para Capa Anterior
    qryGetPrimeiroDocumentoCapa         As rdoQuery           'Obtem o primeiro IdDocto referente a uma determinada Capa
    qryOrdenaCapturaSplitCapa           As rdoQuery           'Ordena campo OrdemCaptura de todos documentos do Split
    qryGeraOcorrenciaDocumento          As rdoQuery           'Devolve documento com ocorrência
    qryAtualizaFGTS                     As rdoQuery           'Atualiza FGTS na complementacao automatica
    qryAtualizaOcorrenciaCapa           As rdoQuery           'Atualiza ocorrência no registro da tabela capa
    qryInsereMotivoExclusao             As rdoQuery           'Insere motivo de exclusão para capa devolvida automaticamente
    Capa                                As tpCapa             'Guarda situação atual de capa (Não sofre alteração de complementação, Somente após complementação efetivada)
    rstModulo                           As rdoResultset       'Resultset utilizado pelo modulo
End Type

Private Type tpLog
    TipoDoctoAlterado                   As Integer
    DocumentoComplementado              As Integer
    ComplementacaoAutomatica            As Integer
    DevolvidoPorDuplicidadeAuto         As Integer
    EnvelopeMaloteComplementado         As Integer
    DocumentoIlegivel                   As Integer
    EnviarVinculoAuto                   As Integer
    EnviarIlegivelAuto                  As Integer
    EnviarConfirmacaoAgConta            As Integer
    SplitCapaInicial                    As Integer
    SplitCapaFinal                      As Integer
    SplitAnteriorInicial                As Integer
    SplitAnteriorFinal                  As Integer
    SistemaDeletaCapaAuto               As Integer
End Type

Private Modulo                          As tpModulo
Private LOG                             As tpLog
Private iFlagRotacao                    As Integer        'Identificador de rotação de imagem
Private sTempo                          As Integer        'Controle de tempo para ativar (Timer) informando Capa ainda em complementação
Private lcontador                       As Long           'Contador de minutos para controle de Timer
Private bSupervisor                     As Boolean        'Identificador de Documento enviado para ilegível
Private bConfirmaAgConta                As Boolean        'Identificador de documento à ser enviado para confirmação de agencia e conta
Private bDuplicidade                    As Boolean        'Variavel carregada na atualização de Documento (Função G_AtualizaCamposDocumento)
Public sCapaOuDocumento                 As String         'Identificador de documento em complementação (C)-Capa (D)-Documento

'Constantes que identificam as colunas do grid de Documentos
Public iColIdDocto                      As Integer
Public iColdCapa                        As Integer
Public iColTpDocto                      As Integer
Public iColLeitura                      As Integer
Public iColFrente                       As Integer
Public iColVerso                        As Integer
Public iColStatus                       As Integer
Public iColValor                        As Integer
Public iColOrdem                        As Integer
Public bFim                             As Boolean

'Variavel utilizada no tratamento de Erro
Dim sPosicaoErro                        As String

'Constantes contendo o Caption inicial do cabeçalho
Const cst_lblDocumento                  As String = "Documento:  "
Const cst_lblCapa                       As String = "Lote/Capa "
Const cst_lblData                       As String = "Data:  "

Private Sub MostrarImagem()

    Dim Ret As Long
    '''''''''''''''''''''''''''''''''
    ' Mostrar documento na LeadTools'
    '''''''''''''''''''''''''''''''''
    With Lead1
       .Visible = False
       On Error Resume Next
       If Me.Lead1.Tag = "0" Then
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Geral.Documento.Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Frente, 0, 0, 1
            End If
        Else
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Geral.Documento.Verso, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Verso, 0, 0, 1
            End If
        End If
       ' se imagem for da ls500, mostra mais escura
       If Geral.Documento.Ordem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for da canon, mostra a 50%
       If Geral.Documento.Ordem <> "1" Then
          .PaintZoomFactor = 100
       Else
          .PaintZoomFactor = 50
       End If
             
      .Visible = True
    End With
    
    'Posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
End Sub

Private Sub Complementar()

On Error GoTo Err_Complementar

    Dim iTipoDocto                  As enumTipoDocto
    Dim bCancelou                   As Boolean
    Dim bAtualizarEnvelope          As Boolean
    Dim iSituacao                   As Integer
    Dim bExecutou                   As Boolean
    Dim sLeitura                    As String
    Dim iDoctoAtual                 As Integer
    Dim bErroDocto                  As Boolean
    Dim bDocumentoAnterior          As Boolean
    Dim bComplementacaoAutomatica   As Boolean
    Dim iUltimoDocto                As Integer
    'Informa se Capa em complementação já foi complementada
    'anteriormente ou se é a primeira vez em complementação
    Dim iStatusDoctoCapa            As Integer
    Dim iErroSplit                  As Integer
    Dim bCapaAnterior               As Boolean
    'Numero de tentativas de obter capa para complementar
    Dim iTentativas                 As Integer
    Dim rst                         As RDO.rdoResultset
    Dim sStr                        As String
    
    hCtl = Lead1.hwnd

    bFim = False
    bCancelou = False
    bDocumentoAnterior = False
    bAtualizarEnvelope = False
    
    'Informa se ocorreu complementação
    grdDocumentos.Tag = False
    
    Do  'Procedimento para tratar Capa
                
        iTentativas = 0
        
        rdoErrors.Clear
        bAtualizarEnvelope = False
        
InicioDeCapa:
        
        'Inicia controle de tempo
        sTempo = 0
        TmrAtualiza.Enabled = False
        
        'Limpa grid contendo dados de documento
        grdDocumentos.Rows = 0
    
        'Inicializa Controle de inversão da imagem e rotacionamento
        ' Lead1.Tag (0)-Sem Inversão (1)-Com Inversão
        Lead1.Tag = "0"
        iFlagRotacao = 0
        
        'Limpa Cabeçalho do Form
        Call Cabecalho(True)
        
        bSupervisor = False
        bCapaAnterior = False
        
        InicializaDocumento
        InicializaCapa
        
        sPosicaoErro = "GetEnvDig"
        Modulo.qryGetIdCapaComplementacao.rdoParameters(1) = Geral.DataProcessamento
        'Intervalo de tempo para busca de capa disponível
        Modulo.qryGetIdCapaComplementacao.rdoParameters(2) = Geral.Intervalo
        Modulo.qryGetIdCapaComplementacao.Execute
        iTentativas = iTentativas + 1

        'Verifica se existe capa a ser complementada
        If Modulo.qryGetIdCapaComplementacao.rdoParameters(0) = 1 Then

            'Tenta no maximo 3 vezes obter capa para complementar
            If iTentativas < 1 Then
                GoTo InicioDeCapa
            Else
                iTentativas = 0
            End If
            
            'Verifica se já ocorreu complementação
            Beep
            If grdDocumentos.Tag Then
                Me.Lead1.Visible = False
                MsgBox "Não existem mais Documentos a serem complementados!", vbInformation + vbOKOnly, App.Title
            End If
            
            'Sai de complementação
            Exit Do
        End If
        
        'Informa que já houve digitação de documentos
        grdDocumentos.Tag = True
        
        'Verifica se ocorreu erro na leitura de capa
        If Modulo.qryGetIdCapaComplementacao.rdoParameters(0) = 2 Then
            Beep
            If MsgBox("Problema na verificação de documento a ser complementado. Nova tentativa", vbExclamation + vbYesNo, App.Title) = vbYes Then
                GoTo InicioDeCapa
            Else
                Exit Do
            End If
        End If
        
        'Carrega variaveis globais com dados da Capa
        Geral.Capa.IdLote = Modulo.qryGetIdCapaComplementacao.rdoParameters("@IdLote").Value
        Geral.Capa.IdCapa = Modulo.qryGetIdCapaComplementacao.rdoParameters("@IdCapa").Value
        'Carrega variaveis com situação atual de capa (Sem alteração de complementação)
        Modulo.Capa.IdLote = Geral.Capa.IdLote
        Modulo.Capa.IdCapa = Geral.Capa.IdCapa
        
        'Apresenta Cabeçalho parcial
        Call Cabecalho(False)
        
        '--------------------------------------------
        '   Obtem dados complementares da capa
        '--------------------------------------------
        sPosicaoErro = "ObtDadosCapa"
        Modulo.qryDadosCapa.rdoParameters(1) = Geral.DataProcessamento
        Modulo.qryDadosCapa.rdoParameters(2) = Geral.Capa.IdCapa
        Set Modulo.rstModulo = Modulo.qryDadosCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        
        'Verifica se ocorreu erro
        If Modulo.qryDadosCapa.rdoParameters(0).Value <> 0 Then
            Beep
            If MsgBox("Não foi possível ler informações de Capa. Continua ", vbExclamation + vbYesNo, App.Title) = vbYes Then
                GoTo InicioDeCapa
            Else
                Exit Do
            End If
        End If
        
        'Inicia controle de tempo para informar Capa ainda em Complementação
        sTempo = 0
        TmrAtualiza.Enabled = True
        
        'Verifica se Documento inicial não refere-se a Capa de Envelope/Malote
        If InStr("EM", Modulo.rstModulo!IdEnv_Mal) = 0 Then
            Geral.Capa.Capa = Modulo.rstModulo!Capa
            
            'Finaliza atualização da Capa devido a mudança de Status
            TmrAtualiza.Enabled = False
            
            Beep
            If AtualizaStatusCapa(Geral.Capa.IdCapa, "5") Then
                MsgBox "Capa ( " & CStr(Geral.Capa.Capa) & " ) enviada para supervisor devido a documento ilegível!", vbOKOnly, App.Title
                
                'Grava Log de ocorrência
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarIlegivelAuto)
                GoTo InicioDeCapa
            Else
                If MsgBox("Não foi possível enviar capa ( " & CStr(Geral.Capa.Capa) & " ) para supervisor. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                    GoTo InicioDeCapa
                Else
                    Exit Do
                End If
            End If
        End If
        
        Geral.Capa.IdEnv_Mal = Modulo.rstModulo!IdEnv_Mal
        Geral.Capa.Capa = Modulo.rstModulo!Capa
        Geral.Capa.Num_Malote = Modulo.rstModulo!Num_Malote
        Geral.Capa.AgOrig = Modulo.rstModulo!AgOrig
        Geral.Capa.Status = Modulo.rstModulo!Status
        Geral.Capa.Duplicidade = Modulo.rstModulo!Duplicidade
        If Modulo.rstModulo!AgOrig <> 0 Then
            If Not CarregaAGENF(Modulo.rstModulo!AgOrig) Then
                Call AtualizaStatusCapa(Geral.Capa.IdCapa, "5")
                MsgBox "Código de Agência não localizado na (AGENF), favor contatar o suporte!" & vbCrLf & vbCrLf & "Capa enviada para Ilegíveis.", vbCritical, App.Title
                Exit Do
            End If
        End If
        
        'Carrega variaveis com situação atual de capa
        Modulo.Capa.IdEnv_Mal = Geral.Capa.IdEnv_Mal
        Modulo.Capa.Capa = Geral.Capa.Capa
        Modulo.Capa.Num_Malote = Geral.Capa.Num_Malote
        Modulo.Capa.AgOrig = Geral.Capa.AgOrig
        Modulo.Capa.Status = Geral.Capa.Status
        Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
        
        Modulo.Capa.agefsdtmvan = Geral.Capa.agefsdtmvan
        Modulo.Capa.agefsdtmvat = Geral.Capa.agefsdtmvat
        Modulo.Capa.agefsestado = Geral.Capa.agefsestado
        Modulo.Capa.agefsstmovi = Geral.Capa.agefsstmovi
        
        'Cabeçalho
        Call Cabecalho(False)
        
        '--------------------------------------------
        '   Obtem documento para complementação
        '--------------------------------------------
        sPosicaoErro = "ObtDoctos"
        Modulo.qryGetTodosDocumentosCapa(1) = Geral.DataProcessamento
        Modulo.qryGetTodosDocumentosCapa(2) = Geral.Capa.IdCapa
        Set Modulo.tbDocumentos = Modulo.qryGetTodosDocumentosCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        
        'Verifica se ocorreu erro
        If Modulo.qryGetTodosDocumentosCapa.rdoParameters(0) <> 0 Then
            Beep
            If MsgBox("Não foi possível ler informações de Documento. Continua ", vbExclamation + vbYesNo, App.Title) = vbYes Then
                GoTo InicioDeCapa
            Else
                Exit Do
            End If
        End If
        
        'Verifica se retornou documentos e Tipo de Documento (Primeiro registro)
        'é diferente de Envelope/Malote

        bErroDocto = False
        If Modulo.tbDocumentos.RowCount < 1 Then
            bErroDocto = True
        Else
            If Modulo.tbDocumentos!TipoDocto <> 1 Then bErroDocto = True
        End If
        If bErroDocto Then
            Beep
            MsgBox "Problema na leitura de documentos devido à falha na captura." + vbCrLf + vbCrLf + _
                    "Capa será deletada pelo sistema." + vbCrLf + vbCrLf + _
                    "Capa Nr. " + CStr(Geral.Capa.Capa), vbCritical + vbOKOnly, App.Title
            'Finaliza atualização da Capa devido a mudança de Status
            TmrAtualiza.Enabled = False
                
            If AtualizaStatusCapa(Geral.Capa.IdCapa, "D") Then
                MsgBox "Capa deletada pelo sistema." + vbCrLf + vbCrLf + _
                        "Capa Nr. " + CStr(Geral.Capa.Capa), vbOKOnly, App.Title
                'Grava ocorrência para capa
                Call GeraOcorrenciaCapa
                'Grava Log de ocorrência
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.SistemaDeletaCapaAuto)
                'Grava Motivo de exclusão para capa sem documentos (Deletada automaticamente)
                Call MotivoExclusao(Geral.Capa.IdCapa, "Capa sem documentos devolvida automaticamente")
                
                GoTo InicioDeCapa
            Else
                If MsgBox("Não foi possível eliminar capa com problema de captura. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                     GoTo InicioDeCapa
                Else
                    Exit Do
                End If
            End If
        End If
        
        'Verifica se Capa já complementada
        iStatusDoctoCapa = Modulo.tbDocumentos!Status
        
        'Carrega Grid com todos documentos
        Do Until Modulo.tbDocumentos.EOF
            grdDocumentos.Rows = grdDocumentos.Rows + 1
            grdDocumentos.Row = (grdDocumentos.Rows - 1)
            grdDocumentos.Col = iColIdDocto:    grdDocumentos.Text = Modulo.tbDocumentos!IdDocto
            grdDocumentos.Col = iColdCapa:      grdDocumentos.Text = Modulo.tbDocumentos!IdCapa
            grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = Modulo.tbDocumentos!TipoDocto
            grdDocumentos.Col = iColFrente:     grdDocumentos.Text = Modulo.tbDocumentos!Frente
            grdDocumentos.Col = iColVerso:      grdDocumentos.Text = Modulo.tbDocumentos!Verso
            grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Modulo.tbDocumentos!Status
            grdDocumentos.Col = iColValor:      grdDocumentos.Text = Modulo.tbDocumentos!Valor
            grdDocumentos.Col = iColOrdem:      grdDocumentos.Text = Modulo.tbDocumentos!Ordem
            
            grdDocumentos.Col = iColLeitura:    sLeitura = Trim(Modulo.tbDocumentos!Leitura)
            grdDocumentos.Text = sLeitura
            
            'Verifica se TODO Campo leitura = "0"
            If Len(sLeitura) > 0 Then
                If sLeitura = String(Len(sLeitura), "0") Then
                    grdDocumentos.Text = ""
                End If
            End If
            Modulo.tbDocumentos.MoveNext
        Loop
        
        'Fecha ResultSet
        Modulo.tbDocumentos.Close
        Modulo.rstModulo.Close
        
        'Vai para o primeiro registro do Documento
        grdDocumentos.Row = 0
        
        'Apresenta Cabeçalho do Form
        Call Cabecalho(False, grdDocumentos.Row + 1, grdDocumentos.Rows)
        
        'Carrega Parte de variáveis globais do Documento
        Call CarregaDocumento
        
        '----------------------------------------------------------------------
        '   Verifica se existe Envelope ou Malote para Complementar
        '----------------------------------------------------------------------
        sCapaOuDocumento = "C"
        If Geral.Capa.IdEnv_Mal = "E" Then
            If (Geral.Capa.AgOrig = 0) Or (Geral.Capa.Capa = 0) Or (Geral.Capa.Capa = 9) Then
                bAtualizarEnvelope = True
            Else
                'Para complementação automática, atualizar estatística por agência e verificar duplicidade (Somente se Capa não complementada)
                If Geral.Documento.Status = "0" Then
                    sPosicaoErro = "ComplEnvAuto"
                    If Not ComplementaCapaAutomatico("E") Then
                        If MsgBox("Não foi possível atualizar documento referente a Capa de Envelope  (Modo Automático)." + vbCrLf + vbCrLf + "Continua ?", vbCritical + vbYesNo, App.Title) = vbYes Then
                            GoTo InicioDeCapa
                        Else
                            Exit Do
                        End If
                    End If
                End If
            End If
        
        ElseIf Geral.Capa.IdEnv_Mal = "M" Then
            If (Geral.Capa.AgOrig = 0) Or (Geral.Capa.Capa = 0) Or _
                (Geral.Capa.Num_Malote = 0) Or (Geral.Capa.Capa = 9) Then
                bAtualizarEnvelope = True
            Else
                'Para complementação automática, atualizar estatistica por agência e verificar duplicidade (Somente se Capa não complementada)
                If Geral.Documento.Status = "0" Then
                    sPosicaoErro = "ComplMaloteAuto"
                    If Not ComplementaCapaAutomatico("M") Then
                        If MsgBox("Não foi possível atualizar documento referente a Capa de Malote  (Modo Automático)." + vbCrLf + vbCrLf + "Continua ?", vbCritical + vbYesNo, App.Title) = vbYes Then
                            GoTo InicioDeCapa
                        Else
                            Exit Do
                        End If
                    End If
                End If
            End If
        End If
        
        If bAtualizarEnvelope Then

            'Apresenta Imagem do Documento
            MostrarImagem
            
            If Geral.Capa.IdEnv_Mal = "E" Then
                'Apresenta verso da imagem para envelope
                cmdFrenteVerso_Click
                Load Envelope
                With Envelope
                    .SetParent Complementacao
                    .SetPosition ((Screen.Width - .Width) / 2), 5600
                    .Show vbModal, Me
                End With
                Unload Envelope
            Else
                Load Malote
                With Malote
                    .SetParent Complementacao
                    .SetPosition ((Screen.Width - .Width) / 2), 5450
                    .Show vbModal, Me
                End With
                Unload Malote
            End If

            'Verifica se cancelou digitação do Envelope/Malote
            If Not IIf(Geral.Capa.IdEnv_Mal = "E", Envelope.Alterou, Malote.Alterou) Then
                
                'Verifica se capa não identificada pela Captura de Documentos
                If Geral.Capa.Capa = 9 Then
CapaIlegivel:
                    Load DocumentoDesconhecido
                    With DocumentoDesconhecido
SelecionaCapa:
                        
                        .SetParent Complementacao
                        .SetPosition ((Screen.Width - .Width) / 2), 5200
                        
                        .Left = 0
                        .Width = Width
                        '-------------------------------------------------------------------
                        'Se Capa já definida e existe capa anterior complementada, permite
                        'passar todos documento desta capa para a última capa complementada
                        'caso operador entre com tipo de documento <> de Capa Envelope/Malote
                        '-------------------------------------------------------------------
                        If Geral.Capa.Capa <> 9 Then
                            .InibirOpcoes (-1)  'Habilita todas opções de documentos na tela de Docto Desconhecido
                            .Show vbModal, Me
                        Else
                            .InibirOpcoes (-9)  'Habilita somente opções de Capa Malote e Envelope na tela de Docto Desconhecido
                            .Show vbModal, Me
                            .InibirOpcoes (-1)  'Habilita todas opções de documentos na tela de Docto Desconhecido
                        End If
                        Unload DocumentoDesconhecido
                        
                        If .Supervisor Then

                            'Finaliza atualização da Capa devido a mudança de Status
                            TmrAtualiza.Enabled = False
                            
                            If AtualizaStatusCapa(Geral.Capa.IdCapa, "5") Then
                            
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.DocumentoIlegivel)
                                
                                '''''''''''''''''''''''''''''''
                                'Verifica se existe Comentario'
                                '''''''''''''''''''''''''''''''
                                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                                
                                sStr = ""
                                If Not rst.EOF() Then
                                    sStr = rst!Comentario
                                End If
                                '''''''''''''''''''''
                                'Insere ControleCapa'
                                '''''''''''''''''''''
                                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    'Se não conseguiu inserir, isto foi somente um detalhe'
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                                End If
                                
                                
                                GoTo InicioDeCapa
                            Else
                                Beep
                                MsgBox "Não foi possível enviar capa para supervisor !", vbCritical + vbOKOnly, App.Title
                                bFim = True
                                Exit Do
                            End If

                        ElseIf .Cancelou Then
                            
                            '---------------------------------------------------------------------
                            'Se Cancelou e documento não foi definido como capa de Malote/Envelope
                            'e capa é <> (9) então abandona complementação
                            '---------------------------------------------------------------------
                            If Geral.Capa.Capa <> 9 Then
                                'Finaliza atualização da Capa devido a mudança de Status
                                TmrAtualiza.Enabled = False
                            
                                'Retorna status de capa para Digitalizada
                                Call AtualizaStatusCapa(Geral.Capa.IdCapa, "1")
                            
                                Lead1.Visible = False
        
                                bFim = True
                                Exit Do
                                    
                            End If
                            
                            '---------------------------------------------------------------------
                            'Se Cancelou e documento não foi definido como capa de Malote/Envelope
                            'e capa é = (9) então abandona envia para Ilegíveis
                            '---------------------------------------------------------------------
'                            Geral.Banco.RollbackTrans
                            'Finaliza atualização da Capa devido a mudança de Status
                            TmrAtualiza.Enabled = False
                            
                            If AtualizaStatusCapa(Geral.Capa.IdCapa, "5") Then
                            
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.DocumentoIlegivel)
                                
                                '''''''''''''''''''''''''''''''
                                'Verifica se existe Comentario'
                                '''''''''''''''''''''''''''''''
                                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                                
                                sStr = ""
                                If Not rst.EOF() Then
                                    sStr = rst!Comentario
                                End If
                                '''''''''''''''''''''
                                'Insere ControleCapa'
                                '''''''''''''''''''''
                                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    'Se não conseguiu inserir, isto foi somente um detalhe'
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                                End If
                                
                                
                                GoTo InicioDeCapa
                            Else
                                Beep
                                MsgBox "Não foi possível enviar capa para supervisor !", vbCritical + vbOKOnly, App.Title
                                bFim = True
                                Exit Do
                            End If
                        Else
                                                        
                            If .TipoDocto = 99 Then     'Capa de Malote
                                Geral.Capa.IdEnv_Mal = "M"
                                
                                Load Malote
                                With Malote
                                    .SetParent Complementacao
                                    .SetPosition ((Screen.Width - .Width) / 2), 5450
                                    .Show vbModal, Me
                                End With
                                Unload Malote
                                If Not Malote.Alterou Then GoTo SelecionaCapa
                                
                            ElseIf .TipoDocto = 1 Then      'Capa de Envelope
                                Geral.Capa.IdEnv_Mal = "E"
                                
                                Load Envelope
                                With Envelope
                                    .SetParent Complementacao
                                    .SetPosition ((Screen.Width - .Width) / 2), 5600
                                    .Show vbModal, Me
                                End With
                                Unload Envelope
                                If Not Envelope.Alterou Then GoTo SelecionaCapa
                            
                            Else    'Documento (Split para Capa Anterior)
                                
                                iErroSplit = SplitAnterior
                                
                                'Verifica se houve erro de leitura
                                If iErroSplit = 1 Then GoTo CapaIlegivel
                                
                                'Verifica se houve alteração ou Erro de execução
                                If iErroSplit = 2 Or iErroSplit = 9 Then
                                    bFim = True
                                    Exit Do
                                End If
                                'Se completou Split, vai para complementação
                                iTipoDocto = .TipoDocto
                                iUltimoDocto = grdDocumentos.Row
                                bCapaAnterior = True
                                GoTo CapaAnterior
                                
                            End If
                        End If
                    
                    End With

                Else
                    '--------------------------------------------------------------------------------
                    'Verifica se continua complementação da Capa já definida ou sai de Complementação
                    '--------------------------------------------------------------------------------

                    If MsgBox("Complementação de Capa cancelada pelo usuário. Continua complementação ", vbExclamation + vbYesNo, App.Title) = vbYes Then
                        GoTo CapaIlegivel
                    Else
                        'Finaliza atualização da Capa devido a mudança de Status
                        TmrAtualiza.Enabled = False
                    
                        'Retorna status de capa para Digitalizada
                        Call AtualizaStatusCapa(Geral.Capa.IdCapa, "1")
                    
                        Lead1.Visible = False

                        bFim = True
                        Exit Do
                    End If
                End If
            End If
            
            'Atualiza Documento para (Status =Complementado), (TpDocto=Envelope)
            'Obs: Capa de Envelope já foi atualizado no form (Envelope)

            If bSupervisor Then     'Documento vai para supervisor somente se Capa = "9" e
                                    'complementação enviada para supervisor
                'Se houve inversão de imagem, atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , 1)
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , 1, , , Geral.Documento.Verso, Geral.Documento.Frente)
                End If
            Else
                'Se houve inversão de imagem, atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , 1, , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , 1, , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If
            End If
                
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                If MsgBox("Não foi possível atualizar documento referente a Capa. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                    GoTo InicioDeCapa
                Else
                    Exit Do
                End If
            End If

            If bSupervisor Then
                'Grava Log de ocorrência
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.DocumentoIlegivel)
                
            Else
                'Grava Log de ocorrência
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnvelopeMaloteComplementado)
            End If
            
            'Verifica se existe a mesma capa com status = 0 para capa não definida
            'pela Vips, se Sim excluir a capa recepcionada somente
'            If Modulo.capa.capa = 9 Then

                sPosicaoErro = "RemoveCapa"
                With Modulo.qryRemoveCapaRecepcionada
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Geral.Capa.Capa
                    .rdoParameters(3) = Geral.Capa.AgOrig
                    .rdoParameters(4) = Geral.Capa.Num_Malote
                    .Execute
                    If .rdoParameters(0).Value <> 0 Then
                        Beep
                        If MsgBox("Não foi possível atualizar documento referente a Capa. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                            GoTo InicioDeCapa
                        Else
                            Exit Do
                        End If
                    End If
                End With
            
            'Atualiza variaveis do modulo com variaveis complementadas
            Modulo.Capa.AgOrig = Geral.Capa.AgOrig
            Modulo.Capa.Capa = Geral.Capa.Capa
            Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
            Modulo.Capa.Status = Geral.Capa.Status
            Modulo.Capa.Num_Malote = Geral.Capa.Num_Malote
            Modulo.Capa.IdEnv_Mal = Geral.Capa.IdEnv_Mal
            
            Modulo.Capa.agefsdtmvan = Geral.Capa.agefsdtmvan
            Modulo.Capa.agefsdtmvat = Geral.Capa.agefsdtmvat
            Modulo.Capa.agefsestado = Geral.Capa.agefsestado
            Modulo.Capa.agefsstmovi = Geral.Capa.agefsstmovi

        Else
            'Atualiza status de Documento referente a capa para (1)-Complementado
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1") Then
                Beep
                If MsgBox("Não foi possível atualizar documento referente a Capa. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                    GoTo InicioDeCapa
                Else
                    Exit Do
                End If
            End If
            
            'Grava Log de ocorrência somente se complementado pela primeira vez
            If iStatusDoctoCapa = 0 Then
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.ComplementacaoAutomatica)
            End If
        End If

CapaAnterior:

        Do  'Procedimento para tratar documentos
            
            'Verifica se Complementação do primeiro documento da Capa Anterior (bCapaAnterior = true)
            If bCapaAnterior Then
                bCapaAnterior = False
            Else

InicioDeDocumento:
            
                'Desabilita visão de Imagem
                Lead1.Visible = False
                                
                If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then
                    'Finaliza atualização da Capa devido a mudança de Status
                    TmrAtualiza.Enabled = False
                    
                    'Atualizar o Status da Capa para forçar o tempo para pegar capa perdida
                    'uma vez que o TmrAtualiza está desligado e não existe begintran
                    Call AtualizaStatusCapa(Geral.Capa.IdCapa, "2")
                    
                    If Not EncerraCapa() Then
                        Beep
                        If MsgBox("Não foi possível encerrar o Envelope/Malote. Nova Tentativa ", vbCritical + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                            GoTo InicioDeDocumento
                        Else
                            MsgBox "Não foi possível Encerrar Envelope/Malote. Verifique ocorrência. ", vbCritical + vbOKOnly, App.Title
                            'Força Saída da complementação
                            bCancelou = False
                            bFim = True
                            Exit Do
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Reset do envio da capa para confirmação de Agencia e Conta'
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    bConfirmaAgConta = False
                    
                    'Sai de Documentos e Inicia Nova Capa de Envelope/Malote
                    Exit Do
                End If
                    
                If Not bDocumentoAnterior Then
                    grdDocumentos.Row = grdDocumentos.Row + 1
                    iUltimoDocto = grdDocumentos.Row
                    
                    'Verifica se documento já complementado
                    grdDocumentos.Col = iColStatus
                    If grdDocumentos.Text <> "0" Then GoTo InicioDeDocumento
    
                    'Se ultima digitação foi Capa de OCT, força  OCT (Verifica antes se realmente docto nâo é de outro tipo)
                    If iTipoDocto = 37 Then
                        If Len(grdDocumentos.Text) > 0 Then
                            grdDocumentos.Col = iColLeitura
                            If IdentificaTipoDocto(grdDocumentos.Text) <> 0 Then iTipoDocto = 0
                        End If
                    End If
                    
                    If iTipoDocto <> 37 Then
                        'Carrega variavel Tipo de Documento
                        grdDocumentos.Col = iColTpDocto
                        iTipoDocto = grdDocumentos.Text
                    End If
                
                    'Inicializa Controle de inversão da imagem e rotacionamento
                    ' Lead1.Tag (0)-Sem Inversão (1)-Com Inversão
                    Lead1.Tag = "0"
                    iFlagRotacao = 0
                    
                    
                    'Se documento não identificado, verifica Tipo de Documento através do campo Leitura
                    If iTipoDocto = 0 Then
                        grdDocumentos.Col = iColLeitura
                        iTipoDocto = IdentificaTipoDocto(grdDocumentos.Text)
                        'Acerta Grid de navegação com o tipo de documento identificado pelo sistema
                        If iTipoDocto <> 0 Then
                            grdDocumentos.Col = iColTpDocto
                            grdDocumentos.Text = iTipoDocto
                    
                            'Força apresentação do verso da imagem de envelope somente na primeira
                            'vez em que a imagem é apresentada para complementação
                            If iTipoDocto = 1 Then
                                If grdDocumentos.Row = iUltimoDocto And Lead1.Tag = "0" Then
                                    cmdFrenteVerso_Click
                                End If
                            End If
                        
                        End If
                    End If
                
                End If
                    
                'Apresenta Cabeçalho do Form
                Call Cabecalho(False, grdDocumentos.Row + 1, grdDocumentos.Rows)
            End If

MostrarNovamente:
                
            'Carrega Situação Atual de Capa e Documento
            Call CarregaCapa
            Call CarregaDocumento

            If bDocumentoAnterior Then
                'Apresenta Cabeçalho do Form
                Call Cabecalho(False, grdDocumentos.Row + 1, grdDocumentos.Rows)
                'Inicializa Controle de inversão da imagem e rotacionamento
                If iTipoDocto = 0 Then
                    Lead1.Tag = "0"
                    iFlagRotacao = 0
                End If
            End If
                
            If iFlagRotacao = 0 Then
                'Apresenta Imagem do Documento
                MostrarImagem
            End If
            
            bCancelou = False

            If bDocumentoAnterior And iTipoDocto <> 0 Then
                bDocumentoAnterior = False
            End If
                
            'Verifica tipo do documento

            Select Case iTipoDocto
                
                Case 999 ' Concessionaria que não necessita de digitação (Arrec. Eletrônica)

                    Call ComplementaConcessionaria(iSituacao)
                    'Se código de barras com duplicidade ou Erro de execução, envia para Arrecadação convencional
                    
                    If iSituacao <> 0 Then
                        If iSituacao = 4 Then
                            iTipoDocto = 8
                        Else
                            iTipoDocto = 27
                        End If
                        bCancelou = True
                        GoTo MostrarNovamente
                    End If
                    
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    
                    If iSituacao = 0 Then
                        'Grava LOG para Tipo de Documento Modificado
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.ComplementacaoAutomatica)
                            End If
                        End If
                    End If
                
                Case 888, 39    'Capa de OCT

                    'Se capa com CMC7, dispensa complementação manual e complementa automaticamente
                    If grdDocumentos.Row = iUltimoDocto And IdentificaTipoDocto(Geral.Documento.Leitura) = 888 Then
                        'Informa a complementação de Capa OCT que não necessita digitação
                        iSituacao = 9
                        Call ComplementaCapaOCT(iSituacao)
                        bComplementacaoAutomatica = True
                    Else
                        Call ComplementaCapaOCT(iSituacao)
                        bComplementacaoAutomatica = False
                    End If

                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    
                    'Se complementado Capa de OCT, Força digitação de OCT
                    If iSituacao = 0 Then
                        'Grava LOG para Tipo de Documento Modificado
                        If iTipoDocto = 888 Then iTipoDocto = 39
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            If bComplementacaoAutomatica Then
                                'Grava Log de ocorrência para complementação automática
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.ComplementacaoAutomatica)
                            Else
                                'Grava Log de ocorrência para complementação manual
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                        
                        'Força digitação de OCT
                        iTipoDocto = 37
                    End If

                Case 99 ' Malote
                    
                    Call ComplementaMalote(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        'Inicia controle de tempo para informar Capa ainda em Complementação
                        sTempo = 0
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                    
                    'Retorna/Inicia o controle de tempo para informar que capa em complementação
                    TmrAtualiza.Enabled = True
                        
                Case 1 ' Envelope
                    
                    Call ComplementaEnvelope(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        'Inicia controle de tempo para informar Capa ainda em Complementação
                        sTempo = 0
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                    
                    'Retorna/Inicia o controle de tempo para informar que capa em complementação
                    TmrAtualiza.Enabled = True
                        
                Case 2, 3 ' Depósito
                    
                    'Se Tipo Docto já foi definido, permanece Tipo de Docto anterior (LOG)
                    grdDocumentos.Col = iColTpDocto
                    If grdDocumentos.Text <> 0 Then iTipoDocto = grdDocumentos.Text
                    
                    Call ComplementaDeposito(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        'Atualiza Tipo de Documento no Grid de navegação (LOG)
                        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = Geral.Documento.TipoDocto
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            bConfirmaAgConta = True
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If

                Case 4 ' ADCC
                    Call ComplementaADCC(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            bConfirmaAgConta = True
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If
                    
                    ' Cheque
                Case 5, etpdocChequeTerceiroPagto, etpdocChequeDeposito
                    
                    'Se Tipo Docto já foi definido, permanece Tipo de Docto anterior (LOG)
                    grdDocumentos.Col = iColTpDocto
                    If grdDocumentos.Text <> 0 Then iTipoDocto = grdDocumentos.Text
                    
                    'Se cheque com valor impresso, dispensa complementação manual e compl. automaticamente
                    If grdDocumentos.Row = iUltimoDocto And Geral.Documento.ValorTotal <> 0 Then
                        'Informa a complementação de cheque que não necessita digitação
                        iSituacao = 9
                        Call ComplementaCheque(iSituacao)
                        bComplementacaoAutomatica = True
                        'Se dados de cheque estão errado, força compl. cheque manual
                        If iSituacao = 1 Then
                            Me.grdDocumentos.Col = iColValor: Me.grdDocumentos.Text = 0
                            GoTo MostrarNovamente
                        End If
                        
                    Else
                        Call ComplementaCheque(iSituacao)
                        bComplementacaoAutomatica = False
                    End If

                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        'Atualiza Tipo de Documento no Grid de navegação (LOG)
                        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = Geral.Documento.TipoDocto
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                If bComplementacaoAutomatica Then
                                    'Grava Log de ocorrência para complementação automática
                                    Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.ComplementacaoAutomatica)
                                Else
                                    'Grava Log de ocorrência para complementação manual
                                    Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                                End If
                            End If
                        End If
                    
                    End If
                        
                    ' Concessionária com Valor expresso em Reais
                Case 8, etpdocAgua, etpdocGas, etpdocLuz, etpdocTelefone
                    
                    'Se Tipo Docto já foi definido, permanece Tipo de Docto anterior (LOG)
                    grdDocumentos.Col = iColTpDocto
                    If grdDocumentos.Text <> 0 Then iTipoDocto = grdDocumentos.Text
                    
                    Call ComplementaArrecEletronica(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        'Atualiza Tipo de Documento no Grid de navegação (LOG)
                        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = Geral.Documento.TipoDocto
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If
                
                    ' Concessionária Com Valor indexado
                Case 9, etpdocTributosMunicipais, etpdocTributosEstaduais, etpdocTributosFederais
                    
                    'Se Tipo Docto já foi definido, permanece Tipo de Docto anterior (LOG)
                    grdDocumentos.Col = iColTpDocto
                    If grdDocumentos.Text <> 0 Then iTipoDocto = grdDocumentos.Text
                    
                    Call ComplemementaArrecValorIndexado(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        'Atualiza Tipo de Documento no Grid de navegação (LOG)
                        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = Geral.Documento.TipoDocto
                        
                        If Geral.Documento.Duplicidade = 1 Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If
                
                Case 10 ' Ficha de Compensação
                    
                    Call ComplementaFichaCompensacao(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If

                Case 11 ' Inss
                
                Case 12 ' Titulos
                    
                    Call ComplementaTitulos(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If

                
                Case 13 ' Cobrança Registrada
                
                    Call ComplementaCobrancaRegistrada(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If

                Case 14 ' Cobrança Especial

                    Call ComplementaCobrancaEspecial(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If

                Case 15 ' Darm
                
                    Call ComplementaDarm(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                
                Case 16 ' Darf Preto
                
                    Call ComplementaDarfPreto(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                
                Case 17 ' Darf Simples
                
                    Call ComplementaDarfSimples(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                
                Case 18 ' Gare
                        
                    Call ComplementaGareICMS(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                
                Case 27 ' Arrecadação Convencional
                    Call ComplementaArrecConvencional(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do

                    If iSituacao = 0 Then
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Se Capa em duplicidade, devolver documento com ocorrência
                            If Geral.Capa.Duplicidade <> 0 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                            Else
                                'Grava Log de ocorrência
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If
                    
                Case 35 ' GPS
                    
                    Call ComplementaGPS(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                
                Case 36 ' Cartao credito avulso
                    
                    Call ComplementaCartaoAvulso(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                    
                Case 37 ' OCT
                
                    Call ComplementaOCT(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    
                    If iSituacao = 0 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            bConfirmaAgConta = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                        End If
                    End If
                    
                    'Se finalizado ou não conseguiu atualizar, muda
                    'tipo de docto para evitar novo docto como OCT
                    If iSituacao <> 2 Then iTipoDocto = 0
                    
                Case 40 ' FGTS
                
                    Call ComplementaFGTS(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    If iSituacao = 0 Or iSituacao = 3 Then
                        
                        'Se Capa em duplicidade, devolver documento com ocorrência
                        If Geral.Capa.Duplicidade <> 0 Then
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                        Else
                            'Grava Log de ocorrência
                            If iSituacao = 3 Then
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.ComplementacaoAutomatica)
                            Else
                                Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                            End If
                        End If
                    End If
                    
                    'Se FGTS inválido, enviar para ilegíveis
                    If iSituacao = 4 Then
                        ' Se Documento enviado para supervisor, permanece status (0)Complementar
                        ' e Capa de Envelope/Malote vai para Supervisor (5)legíveis
                        Geral.Capa.Status = "5"
                        bSupervisor = True
    
                        'Marca somente Grid de Navegação com status para Supervisor,
                        'na tabela documento permanece como em complementação (Status=0)
                        grdDocumentos.Col = iColStatus
                        grdDocumentos.Text = "D"
                        bDocumentoAnterior = False
                    
                        'Grava Log de ocorrência
                        Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.EnviarIlegivelAuto)
                    End If
                    
                    
                Case 41 ' Lancamento Interno
                    
                    Call ComplementaLancamentoInterno(iSituacao)
                    If iSituacao = 1 Then bCancelou = True
                    'Se Situação = 2 (Erro SP) Apresenta novamente
                    If iSituacao = 2 Then bCancelou = True: Exit Do
                    
                    If iSituacao = 0 Then
                        
                        If Geral.Documento.Status = "D" Then
                            bSupervisor = True
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                         Else
                              'Se Capa em duplicidade, devolver documento com ocorrência
                              If Geral.Capa.Duplicidade <> 0 Then
                                  bSupervisor = True
                                  Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DevolvidoPorDuplicidadeAuto)
                              Else
                                  'Grava Log de ocorrência
                                  Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoComplementado)
                              End If
                         End If
                    End If
                
                Case Else ' Desconhecido
                    
                    Load DocumentoDesconhecido
                    With DocumentoDesconhecido
                        .SetParent Complementacao
                        .SetPosition ((Screen.Width - .Width) / 2), 5200

                        .Left = 0
                        .Width = Width
                        .Show vbModal, Me
                        
                        If .Supervisor Then
                            ' Se Documento enviado para supervisor, permanece status (0)Complementar
                            ' e Capa de Envelope/Malote vai para Supervisor (5)legíveis
                            Geral.Capa.Status = "5"
                            bSupervisor = True

                            'Marca somente Grid de Navegação com status para Supervisor,
                            'na tabela documento permanece como em complementação (Status=0)
                            grdDocumentos.Col = iColStatus
                            grdDocumentos.Text = "D"
                            bDocumentoAnterior = False
                        
                            'Grava Log de ocorrência
                            Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.DocumentoIlegivel)
                        
                        ElseIf .Cancelou Then
                            Unload DocumentoDesconhecido
                            bFim = True
                            bCancelou = True
                            Exit Do
                        
                        ElseIf .DocumentoAnterior Then
                            bDocumentoAnterior = True
                            If Me.grdDocumentos.Row = 0 Then
                                bDocumentoAnterior = False
                                GoTo InicioDeDocumento
                            End If
                            
                            iTipoDocto = etpdocDesconhecido
                            GoTo MostrarNovamente
                            
                        ElseIf iTipoDocto <> .TipoDocto Then

                            ' Setar novo tipo de documento '
                            iTipoDocto = .TipoDocto
                            If iTipoDocto = etpdocEnvelope Then
                                'Inverte Imagem para Envelope
                                If Me.Lead1.Tag = "0" Then cmdFrenteVerso_Click
                            End If
                            
                            GoTo MostrarNovamente
                        Else
                            bDocumentoAnterior = False
                            'Verifica se Documento escolhido é Envelope
                            If iTipoDocto = etpdocEnvelope Then
                                'Inverte Imagem para Envelope
                                If Me.Lead1.Tag = "0" Then cmdFrenteVerso_Click
                            End If
                            
                            GoTo InicioDeDocumento
                        End If
                    End With
            End Select

            ' Se a alteração de documento anterior foi cancelada, retorna ao ultimo
            ' documento em complementação
            If bCancelou And grdDocumentos.Row <> iUltimoDocto Then
                DocumentoDesconhecido.InibirOpcoes (-1) 'Habilita todas opções da pasta Tipo de Documentos

            ElseIf bCancelou Then
                ' Se foi cancelada, colocar documento como '
                ' desconhecido e mostrar tela de escolha   '
                iTipoDocto = etpdocDesconhecido
                GoTo MostrarNovamente
            
            ElseIf grdDocumentos.Row <> iUltimoDocto Then
                DocumentoDesconhecido.InibirOpcoes (-1) 'Habilita todas opções da pasta Tipo de Documentos
            End If
        Loop

        'Finalizou digitação
        If bFim Then
            If bCancelou Then
                'Finaliza atualização da Capa devido a mudança de Status
                TmrAtualiza.Enabled = False

                'Se cancelada digitação, Volta status de Capa para Digitalizada
                If Geral.Capa.Status = "5" Or bSupervisor Then
                    'Marca capa para Ilegíveis devido a algum Docto enviado para supervisor
                    Call AtualizaStatusCapa(Geral.Capa.IdCapa, "5")
                    'Grava Log de ocorrência
                    If Modulo.Capa.Status <> "5" Then
                        Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.EnviarIlegivelAuto)
                    End If
                    
                    '''''''''''''''''''''''''''''''
                    'Verifica se existe Comentario'
                    '''''''''''''''''''''''''''''''
                    Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                    
                    sStr = ""
                    If Not rst.EOF() Then
                        sStr = rst!Comentario
                    End If
                    '''''''''''''''''''''
                    'Insere ControleCapa'
                    '''''''''''''''''''''
                    If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Se não conseguiu inserir, isto foi somente um detalhe'
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                    End If
                Else
                    'Retorna status de capa para Digitalizada
                    Call AtualizaStatusCapa(Geral.Capa.IdCapa, "1")
                End If
            End If
            'Vai para próxima capa de Envelope/Malote
            Exit Do
        End If
    Loop

    'Finaliza controle de tempo para informar que capa continua em complementação
    TmrAtualiza.Enabled = False

    'Limpa Cabeçalho do Form
    Call Cabecalho(True)

    'Desabilita LEAD
    If Not Lead1.Visible Then Lead1.Visible = False

    Exit Sub

Err_Complementar:

    Select Case TratamentoErro("1. Não foi possível trazer o documento para digitação.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbCancel
            iTipoDocto = etpdocDesconhecido
            GoTo InicioDeCapa
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub CmdFechar_Click()


    bFim = True

    ''''''''''''''''''''''''''''''''''''''
    'Loga a acao de Fim Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''
    GravaLog 0, 0, 251


End Sub

Private Sub Form_Activate()

   Call AtualizaAtividade(8)
   
   Me.Top = 0
   Me.Left = 0

   'Inicia processo de Digitação/Espera de documentos para complementação
   Call IniciarComplementacao

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = &H1B Then
        bFim = True
        CmdSair_Click
    End If

End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim Ret As Long
    
    hCtl = Complementacao.Lead1.hwnd
    
    Select Case KeyCode
        Case vbKeyF10
            KeyCode = 0
        Case vbKeyDown
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEDOWN, 0)
        Case vbKeyUp
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEUP, 0)
        Case vbKeyLeft
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEUP, 0)
        Case vbKeyRight
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEDOWN, 0)
        Case vbKeyPageUp
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_PAGEUP, 0)
        Case vbKeyPageDown
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_PAGEDOWN, 0)
        Case vbKeyHome
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
        Case vbKeyEnd
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_BOTTOM, 0)
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_BOTTOM, 0)
    End Select
    
End Sub
Private Sub Form_Load()

    'Constantes que identificam as colunas do grid de Documentos
    iColIdDocto = 0:    iColdCapa = 1:  iColTpDocto = 2:    iColLeitura = 3
    iColFrente = 4:     iColVerso = 5:  iColStatus = 6:     iColValor = 7
    iColOrdem = 8

    'Código de Ação da tabela de LOG
    LOG.TipoDoctoAlterado = 10
    LOG.DocumentoComplementado = 11
    LOG.ComplementacaoAutomatica = 12
    LOG.DevolvidoPorDuplicidadeAuto = 13
    LOG.EnvelopeMaloteComplementado = 14
    LOG.DocumentoIlegivel = 15
    LOG.EnviarVinculoAuto = 16
    LOG.EnviarIlegivelAuto = 17
    LOG.SistemaDeletaCapaAuto = 19
    LOG.SplitCapaInicial = 150
    LOG.SplitCapaFinal = 151
    LOG.SplitAnteriorInicial = 152
    LOG.SplitAnteriorFinal = 153
    LOG.EnviarConfirmacaoAgConta = 18
    
    Screen.MousePointer = vbHourglass
    
    'para o deskew da lead tools
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With
    
    'Inicializar todas query's
    InicializarQuery
    
    'Limpa Cabeçalho do Form
    Call Cabecalho(True)
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    Call GravaLog(0, 0, 160)
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    With Modulo
        .qryDadosCapa.Close
        .qryGetIdCapaComplementacao.Close
        .qryGetTodosDocumentosCapa.Close
        .qryChecarEnvelope.Close
        .qryAtualizaDuplicidadeCapa.Close
        .qryInsereNovaCapa.Close
        .qryAtualizaArrecEletronica.Close
        .qryGetChequeDuplicado.Close
        .qryAtualizaDocumentoExcluido.Close
        .qryAtualizaCheque.Close
        .qryRemoveCapaRecepcionada.Close
        .qryGetIdDocto.Close
        .qrySplitCapaAnterior.Close
        .qryGetPrimeiroDocumentoCapa.Close
        .qryOrdenaCapturaSplitCapa.Close
        .qryGeraOcorrenciaDocumento.Close
        .qryAtualizaOcorrenciaCapa.Close
        .qryInsereMotivoExclusao.Close
    End With
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    Call GravaLog(0, 0, 161)

    
End Sub

Private Sub InicializarQuery()

    With Modulo
        'Seleciona uma capa para complementação
        Set .qryGetIdCapaComplementacao = Geral.Banco.CreateQuery("", "{? = call GetIdCapaComplementacao(?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Intervalo
            .qryGetIdCapaComplementacao.rdoParameters(0).Direction = rdParamReturnValue
            .qryGetIdCapaComplementacao.rdoParameters(3).Direction = rdParamOutput     'IdLote
            .qryGetIdCapaComplementacao.rdoParameters(4).Direction = rdParamOutput     'IdCapa
        
        'Seleciona todos documentos referentes a uma Capa
        Set .qryGetTodosDocumentosCapa = Geral.Banco.CreateQuery("", "{? = call GetTodosDocumentosCapa(?,?)}")
            .qryGetTodosDocumentosCapa.rdoParameters(0).Direction = rdParamReturnValue
        
        'Seleciona dados de capa
        Set .qryDadosCapa = Geral.Banco.CreateQuery("", "{? = call GetCapa(?,?)}")
            'Parâmetros (1)-Data (2)-IdCapa
            .qryDadosCapa.rdoParameters(0).Direction = rdParamReturnValue
        
        'Retorna quantos números de capa existem com a mesma agencia na tabela capa
        Set .qryChecarEnvelope = Geral.Banco.CreateQuery("", "{? = call ChecarCapaEnvelope  (?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-Agencia (3)-Nr Capa (4)-Numero de Registros encontrados (5)-IdCapa
            .qryChecarEnvelope.rdoParameters(0).Direction = rdParamReturnValue
            'Número de capas existentes na tabela CAPA
            .qryChecarEnvelope.rdoParameters(4).Direction = rdParamOutput
            
        'Atualiza Campo Duplicidade e Status da tabela Capa
        Set .qryAtualizaDuplicidadeCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaDuplicidadeCapa (?,?,?,?)}")
            .qryAtualizaDuplicidadeCapa.rdoParameters(0).Direction = rdParamReturnValue
            
        'Insere nova capa de Envelope/Malote
        Set .qryInsereNovaCapa = Geral.Banco.CreateQuery("", "{? = call InsereCapaSplit (?,?,?,?,?,?,?,?)}")
            'Parâmetros (1)-Data (2)-IdLote (3)-IdEnv_Mal (4)-Capa (5)-AgOrig (6)-Num_Malote (7)-IdCapa (8)-Duplicidade
            .qryInsereNovaCapa.rdoParameters(0).Direction = rdParamReturnValue
            .qryInsereNovaCapa.rdoParameters(7).Direction = rdParamOutput
            .qryInsereNovaCapa.rdoParameters(8).Direction = rdParamOutput
        
        'Insere dados de Arrecadação
        Set .qryAtualizaArrecEletronica = Geral.Banco.CreateQuery("", "{? = call AtuArrecEletronicaAutomatica (?,?,?,?,?,?,?,?)}")
            .qryAtualizaArrecEletronica.rdoParameters(0).Direction = rdParamReturnValue
            
        'Verifica duplicidade da numeração de cheque
        Set .qryGetChequeDuplicado = Geral.Banco.CreateQuery("", "{? = call GetChequeDuplicado (?,?)}")
            .qryGetChequeDuplicado.rdoParameters(0).Direction = rdParamReturnValue
        
        'Move documento cheque para duplicado
         Set .qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoExcluido (?,?,?,?,?)}")
            .qryAtualizaDocumentoExcluido.rdoParameters(0).Direction = rdParamReturnValue
        
        'Insere dados de cheque na tabela cheque
        Set .qryAtualizaCheque = Geral.Banco.CreateQuery("", "{? = call AtualizaCheque (?,?,?,?,?,?)}")
            .qryAtualizaCheque.rdoParameters(0).Direction = rdParamReturnValue

        'Remove Capa apenas Recepcionada (Status = 0)
        Set .qryRemoveCapaRecepcionada = Geral.Banco.CreateQuery("", "{? = call RemoveCapaRecepcionada (?,?,?,?)}")
        'Parâmetros (1)-Data (2)-Capa (3)-AgOrig (4)-Num_Malote
            .qryRemoveCapaRecepcionada.rdoParameters(0).Direction = rdParamReturnValue
        
        'Obtem dados de documento
        Set .qryGetIdDocto = Geral.Banco.CreateQuery("", "{call GetIdDocto(?,?)}")
            'Parâmetros (0)-Data (1)-IdDocto

        'Split para Capa Anterior
        Set .qrySplitCapaAnterior = Geral.Banco.CreateQuery("", "{? = call SplitCapaAnterior(?,?,?)}")
            'Parâmetros (1)-Data (2)-IdCapaAnterior (3)-IdCapaAtual
            .qrySplitCapaAnterior.rdoParameters(0).Direction = rdParamReturnValue

        'Obtem o primeiro IdDocto referente a uma determinada Capa
        Set .qryGetPrimeiroDocumentoCapa = Geral.Banco.CreateQuery("", "{? = call GetPrimeiroDocumentoCapa(?,?,?)}")
            .qryGetPrimeiroDocumentoCapa.rdoParameters(0).Direction = rdParamReturnValue
            .qryGetPrimeiroDocumentoCapa.rdoParameters(3).Direction = rdParamOutput
            
        'Ordena campo OrdemCaptura de todos documentos do Split
        Set .qryOrdenaCapturaSplitCapa = Geral.Banco.CreateQuery("", "{? = call OrdenaCapturaSplitCapa(?,?)}")
            .qryOrdenaCapturaSplitCapa.rdoParameters(0).Direction = rdParamReturnValue

        'Devolve documento com ocorrência
        Set .qryGeraOcorrenciaDocumento = Geral.Banco.CreateQuery("", "{? = call GeraOcorrenciaDocumento(?,?)}")
            'Parâmetros (0)-Data (1)-IdDocto
            .qryGeraOcorrenciaDocumento.rdoParameters(0).Direction = rdParamReturnValue
    
        'Atualiza ocorrência de capa
        Set .qryAtualizaOcorrenciaCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrenciaCapa(?,?,?)}")
            'Parâmetros (1)-Data (2)-IdCapa (3)-Ocorrência
            .qryAtualizaOcorrenciaCapa.rdoParameters(0).Direction = rdParamReturnValue
    
        'Atualiza Motivo de exclusão para capa devolvida automaticamente
        Set .qryInsereMotivoExclusao = Geral.Banco.CreateQuery("", "{? = call InsereMotivoExclusao (?,?,?)}")
            .qryInsereMotivoExclusao.rdoParameters(0).Direction = rdParamReturnValue
            
    End With
End Sub

Private Sub InicializaDocumento()
    
    With Geral.Documento
        .IdCapa = 0:            .IdDocto = 0
        .Leitura = "":          .Frente = ""
        .Verso = "":            .Status = ""
        .Alcada = "":           .Autenticado = ""
        .Ocorrencia = 0:        .OcorrenciaOK = ""
        .Ordem = "":            .ValorTotal = 0
        .NSU = "":              .Terminal = 0
        .Vinculo = 0:           .CMC7Associado = ""
        .Duplicidade = 0:       .TipoDocto = 0
        .Atualizacao = 0:       .Transacao = 0
        .Efetivado = False:     .PagtoTerceiro = ""
        .TotalVinculado = 0:    .Excluido = False
        .AjusteInterno = False: .Agencia = 0
        .Conta = 0:             .AgenciaVinculo = 0
        .ContaVinculo = 0
    End With
End Sub

Private Sub CarregaDocumento()
    
    With Geral.Documento
        
        grdDocumentos.Col = iColIdDocto:    .IdDocto = grdDocumentos.Text
        grdDocumentos.Col = iColdCapa:      .IdCapa = grdDocumentos.Text
        grdDocumentos.Col = iColTpDocto:    .TipoDocto = grdDocumentos.Text
        grdDocumentos.Col = iColLeitura:    .Leitura = grdDocumentos.Text
        grdDocumentos.Col = iColFrente:     .Frente = grdDocumentos.Text
        grdDocumentos.Col = iColVerso:      .Verso = grdDocumentos.Text
        grdDocumentos.Col = iColStatus:     .Status = grdDocumentos.Text
        grdDocumentos.Col = iColValor:      .ValorTotal = grdDocumentos.Text
        grdDocumentos.Col = iColOrdem:      .Ordem = grdDocumentos.Text
                                            .Agencia = Geral.Capa.AgOrig
        
    End With

End Sub
Private Sub Cabecalho(Optional bApenasLimpa As Boolean = False, Optional iDoctoDigitados As Integer, Optional iDoctoDigitar As Integer)
        
    If bApenasLimpa Then
        'Limpa Header do formulário
        lblDocumento.Caption = cst_lblDocumento
        lblCapa.Caption = cst_lblCapa
    Else
        'Preenche Header do formulário
        lblDocumento.Caption = cst_lblDocumento & Format(iDoctoDigitados, "#,##0") & "/" & Format(iDoctoDigitar, "#,##0")
        If Geral.Capa.Capa = 0 Then
            lblCapa.Caption = cst_lblCapa & IIf(Geral.Capa.IdEnv_Mal = "M", "de Malote", "de Envelope") & ":  "
        Else
            lblCapa.Caption = cst_lblCapa & Switch(Geral.Capa.IdEnv_Mal = "M", "de Malote", Geral.Capa.IdEnv_Mal = "E", "de Envelope", Geral.Capa.IdEnv_Mal = "", "") & ":  " & Format(Geral.Capa.IdLote, "0000-00000") & " / " & Geral.Capa.Capa
        End If
    End If
   lblData = cst_lblData & Right(Geral.DataProcessamento, 2) & "/" & Mid(Geral.DataProcessamento, 5, 2) & "/" & Left(Geral.DataProcessamento, 4)
    
End Sub

Function SplitCapa() As Boolean
    
Dim iPosicaoAtualGrid As Integer, lNovaIdCapa As Long, lIdDocto As Long

Dim bExecutou As Boolean
Dim lAnteriorIdCapa As Long

On Error GoTo Err_Split

    'Guarda posição atual de Documento no grid
    iPosicaoAtualGrid = grdDocumentos.Row

    SplitCapa = False
    
    'Guarda IdCapa anterior para controle do LOG
    lAnteriorIdCapa = Geral.Capa.IdCapa
    
    'Insere nova capa de envelope/malote tornando-o em complementação (Status = 2)
    With Modulo.qryInsereNovaCapa
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Capa.IdLote       'Permanece o mesmo Número de Lote do envelope anterior
        .rdoParameters(3) = Geral.Capa.IdEnv_Mal    'Indicador de (E)Envelope ou (M)Malote
        .rdoParameters(4) = Geral.Capa.Capa         'Número de Capa
        .rdoParameters(5) = Geral.Capa.AgOrig       'Agencia Origem
        .rdoParameters(6) = Geral.Capa.Num_Malote   'Número de Malote
        .Execute
        
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value <> 0 Then GoTo Exit_Split
        lNovaIdCapa = .rdoParameters("@IdCapa")
        'Atualiza variaveis com dados da nova capa
        Geral.Capa.IdCapa = lNovaIdCapa
        Geral.Capa.Duplicidade = .rdoParameters("@Duplicidade")
        Geral.Capa.Status = "2"
    End With
    
    'Obtem número do Iddocto referente ao documento contendo dados de capa
    grdDocumentos.Col = iColIdDocto
    lIdDocto = grdDocumentos.Text
    
    'Altera IDCapa para todos documento a partir do Novo número de Envelope
    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
    Do
        If Not G_AtualizaCamposDocumento(bDuplicidade, grdDocumentos.Text, lNovaIdCapa) Then GoTo Exit_Split
        
        If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then Exit Do
        'Pega próximo IdDocto para atualização do Novo IdCapa
        grdDocumentos.Row = grdDocumentos.Row + 1
    Loop
    
    'Ordena OrdemCaptura de todos documento referentes à nova Capa
    With Modulo.qryOrdenaCapturaSplitCapa
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = lNovaIdCapa             'IdCapa da Nova Capa gerada por Split
        .Execute
        
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value <> 0 Then GoTo Exit_Split

    End With
    
    'Atualiza o Tipo de Documento para (1)-Envelope/Malote e
    'Leitura e status (Documento referentes a Capa)
    If Geral.Capa.IdEnv_Mal = "E" Then
        'Verificar, se houve inversão de imagem atualiza documento
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        If Lead1.Tag = "0" Then
            bExecutou = G_AtualizaCamposDocumento(bDuplicidade, lIdDocto, , 1, Format(Geral.Capa.Capa, "00000000"), "1")
        Else
            'Altera Imagem Frente/Verso
            bExecutou = G_AtualizaCamposDocumento(bDuplicidade, lIdDocto, , 1, Format(Geral.Capa.Capa, "00000000"), "1", Geral.Documento.Verso, Geral.Documento.Frente)
        End If

    Else
        'Verificar, se houve inversão de imagem atualiza documento
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        If Lead1.Tag = "0" Then
            bExecutou = G_AtualizaCamposDocumento(bDuplicidade, lIdDocto, , 1, Geral.Documento.Leitura, "1")
        Else
            'Altera Imagem Frente/Verso
            bExecutou = G_AtualizaCamposDocumento(bDuplicidade, lIdDocto, , 1, Geral.Documento.Leitura, "1", Geral.Documento.Verso, Geral.Documento.Frente)
        End If

    End If
    
    'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
    If bExecutou And iFlagRotacao <> 0 Then
        bExecutou = AlteraRotacao
    End If
    
    If Not bExecutou Then GoTo Exit_Split
    
    'Grava Log de ocorrência
    Call GravaLog(lAnteriorIdCapa, 0, LOG.SplitCapaInicial)
    Call GravaLog(lNovaIdCapa, 0, LOG.SplitCapaFinal)
    
    Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnvelopeMaloteComplementado)
    
    bConfirmaAgConta = False
    SplitCapa = True
    
Exit_Split:
    'Retorna Grid de navegação para primeiro documento do Novo Envelope/Malote
    grdDocumentos.Row = iPosicaoAtualGrid
    
    Exit Function

Err_Split:
    
    Select Case TratamentoErro("Não foi possível realizar o Split deste Envelope/Malote!", Err, rdoErrors, False)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    GoTo Exit_Split
    
End Function

Private Function EncerraCapa() As Boolean

    Dim sStatusCapa     As String
    Dim rst             As RDO.rdoResultset
    Dim sStr            As String

EncerraCapa = False

On Error GoTo Err_EncerraCapa

    If Geral.Capa.IdEnv_Mal = "E" Then
        '---------------------------------------
        '----           ENVELOPE            ----
        '---------------------------------------
        'Se existe Capa com duplicidade, Manda para Ilegíveis
        If Modulo.Capa.Duplicidade <> 0 Then
            'Atualiza Status e Duplicidade da tabela Capa
            With Modulo.qryAtualizaDuplicidadeCapa
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Modulo.Capa.IdCapa
                .rdoParameters(3) = "D"                     'Capa Devolvida pelo Sistema
                .rdoParameters(4) = 1                       'Identificador de envelope em duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value <> 0 Then GoTo Exit_EncerraCapa
                Geral.Capa.Status = "5"
                
                '''''''''''''''''''''''''''''''
                'Verifica se existe Comentario'
                '''''''''''''''''''''''''''''''
                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                
                sStr = ""
                If Not rst.EOF() Then
                    sStr = rst!Comentario
                End If
                '''''''''''''''''''''
                'Insere ControleCapa'
                '''''''''''''''''''''
                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Se não conseguiu inserir, isto foi somente um detalhe'
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                End If
                
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.DevolvidoPorDuplicidadeAuto)
            
            End With
        Else
            'Atualiza Status e Duplicidade de tabela Capa se este não mais encontra-se com duplicade
            With Modulo.qryAtualizaDuplicidadeCapa
            
                sStatusCapa = IIf(Geral.Capa.Status = "5" Or bSupervisor, "5", "8")
                ''''''''''''''''''''''''''''''''''
                'Verifica se o usuario é terceiro'
                ''''''''''''''''''''''''''''''''''
                If GrupoUsuario(Geral.Usuario, eG_TERCEIRO) Then
                    If sStatusCapa = "8" And bConfirmaAgConta Then
                        sStatusCapa = "L"
                    End If
                End If
            
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Modulo.Capa.IdCapa
                .rdoParameters(3) = sStatusCapa                'Envia capa para Ilegiveis ou Vínculo Automático
                .rdoParameters(4) = 0                          'Identificador de envelope sem duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value <> 0 Then GoTo Exit_EncerraCapa
                Geral.Capa.Status = IIf(Geral.Capa.Status = "5" Or bSupervisor, "5", IIf(sStatusCapa = "L", "L", "8"))
                
                'Verifica se Capa enviada para Ilegíveis
                If Geral.Capa.Status = "5" Then
                    'Verifica se Capa enviada para Ilegíveis, Se status = 5 significa que Gerado Log
                    If Modulo.Capa.Status <> "5" Then
                        Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarIlegivelAuto)
                    End If
                ElseIf Geral.Capa.Status = "8" Then
                    Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarVinculoAuto)
                ElseIf Geral.Capa.Status = "L" Then
                    Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarConfirmacaoAgConta)
                End If
                
                '''''''''''''''''''''''''''''''
                'Verifica se existe Comentario'
                '''''''''''''''''''''''''''''''
                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                
                sStr = ""
                If Not rst.EOF() Then
                    sStr = rst!Comentario
                End If
                '''''''''''''''''''''
                'Insere ControleCapa'
                '''''''''''''''''''''
                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Se não conseguiu inserir, isto foi somente um detalhe'
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                End If
            End With
        End If
    Else
        '---------------------------------------
        '----           MALOTE              ----
        '---------------------------------------
        'Se existe Capa com duplicidade, Manda para Ilegíveis
        If Modulo.Capa.Duplicidade <> 0 Then
            
            'Atualiza Status e Duplicidade da tabela Capa
            With Modulo.qryAtualizaDuplicidadeCapa
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Modulo.Capa.IdCapa
                .rdoParameters(3) = "D"                     'Capa Devolvida pelo Sistema
                .rdoParameters(4) = 1                       'Identificador de Malote em duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value <> 0 Then GoTo Exit_EncerraCapa
                Geral.Capa.Status = "5"
                Call GravaLog(Geral.Capa.IdCapa, 0, LOG.DevolvidoPorDuplicidadeAuto)
                '''''''''''''''''''''''''''''''
                'Verifica se existe Comentario'
                '''''''''''''''''''''''''''''''
                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                
                sStr = ""
                If Not rst.EOF() Then
                    sStr = rst!Comentario
                End If
                '''''''''''''''''''''
                'Insere ControleCapa'
                '''''''''''''''''''''
                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Se não conseguiu inserir, isto foi somente um detalhe'
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                End If
            End With
        Else
            'Atualiza Status e Duplicidade de tabela Capa se este não mais encontra-se com duplicade
            With Modulo.qryAtualizaDuplicidadeCapa
            
                sStatusCapa = IIf(Geral.Capa.Status = "5" Or bSupervisor, "5", "8")
            
                ''''''''''''''''''''''''''''''''''
                'Verifica se o usuario é terceiro'
                ''''''''''''''''''''''''''''''''''
                If GrupoUsuario(Geral.Usuario, eG_TERCEIRO) Then
                    If sStatusCapa = "8" And bConfirmaAgConta Then
                        sStatusCapa = "L"
                    End If
                End If
            
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Modulo.Capa.IdCapa
                .rdoParameters(3) = sStatusCapa                                             'Envia capa para Ilegiveis ou Vínculo Automático
                .rdoParameters(4) = 0                                                       'Identificador de envelope sem duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value <> 0 Then GoTo Exit_EncerraCapa
                Geral.Capa.Status = IIf(Geral.Capa.Status = "5" Or bSupervisor, "5", IIf(sStatusCapa = "L", "L", "8"))
                
                'Verifica se Capa enviada para Ilegíveis
                If Geral.Capa.Status = "5" Then
                    'Verifica se Capa enviada para Ilegíveis, Se status = 5 significa que Gerado Log
                    If Modulo.Capa.Status <> "5" Then
                        Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarIlegivelAuto)
                    End If
                ElseIf Geral.Capa.Status = "8" Then
                    Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarVinculoAuto)
                ElseIf Geral.Capa.Status = "L" Then
                    Call GravaLog(Geral.Capa.IdCapa, 0, LOG.EnviarConfirmacaoAgConta)
                End If
                
                '''''''''''''''''''''''''''''''
                'Verifica se existe Comentario'
                '''''''''''''''''''''''''''''''
                Set rst = GetControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa)
                
                sStr = ""
                If Not rst.EOF() Then
                    sStr = rst!Comentario
                End If
                '''''''''''''''''''''
                'Insere ControleCapa'
                '''''''''''''''''''''
                If Not InsereControleCapa(Geral.DataProcessamento, Geral.Capa.IdCapa, sStr, 8) Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Se não conseguiu inserir, isto foi somente um detalhe'
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    MsgBox "Erro ao inserir ControleCapa.", vbExclamation
                End If
            End With
        End If
    
    End If
    
    EncerraCapa = True
    
'Sai da função
Exit_EncerraCapa:

    Exit Function
    
Err_EncerraCapa:
    
    Select Case TratamentoErro("Não foi possível realizar o encerramento deste Envelope/Malote!", Err, rdoErrors, False)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Sub InicializaCapa()

    With Modulo.Capa
        .AgOrig = 0
        .Capa = 0
        .Duplicidade = 0
        .IdCapa = 0
        .IdEnv_Mal = ""
        .IdLote = 0
        .Num_Malote = 0
        .Status = ""
        .agefsdtmvan = 0
        .agefsdtmvat = 0
        .agefsestado = ""
        .agefsstmovi = 0
    End With
    
    With Geral.Capa
        .agefsdtmvan = 0
        .agefsdtmvat = 0
        .agefsestado = ""
        .agefsstmovi = 0
    End With
    
End Sub
Private Sub CarregaCapa()
    
    Geral.Capa.AgOrig = Modulo.Capa.AgOrig
    Geral.Capa.Capa = Modulo.Capa.Capa
    Geral.Capa.Duplicidade = Modulo.Capa.Duplicidade
    Geral.Capa.Status = IIf(bSupervisor, "5", Modulo.Capa.Status)
    Geral.Capa.IdCapa = Modulo.Capa.IdCapa
    Geral.Capa.IdLote = Modulo.Capa.IdLote
    Geral.Capa.IdEnv_Mal = Modulo.Capa.IdEnv_Mal
    Geral.Capa.Num_Malote = Modulo.Capa.Num_Malote
    
    Geral.Capa.agefsdtmvan = Modulo.Capa.agefsdtmvan
    Geral.Capa.agefsdtmvat = Modulo.Capa.agefsdtmvat
    Geral.Capa.agefsestado = Modulo.Capa.agefsestado
    Geral.Capa.agefsstmovi = Modulo.Capa.agefsstmovi

End Sub
Private Function ComplementaMalote(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação da Capa de Malote
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------
Dim iPosicaoAtualGrid As Integer
sPosicaoErro = "ComplMalote"
Dim bDuplicidadeCapa As Boolean

On Error GoTo Err_ComplementaMalote
    
    sCapaOuDocumento = "D"
    
    'Carrega variaveis de Capa para complementar
    Geral.Capa.Capa = IIf(Len(Trim(Geral.Documento.Leitura)) <> 8, 0, Geral.Documento.Leitura)
    Geral.Capa.Duplicidade = 0
    Geral.Capa.Num_Malote = 0
    bDuplicidadeCapa = False
    
    Load Malote
    With Malote
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5230
        .Show vbModal, Me
        
        'Fecha Form
        Unload Malote
        
        If Not .Alterou Then
            iSituacao = 1: Exit Function
        End If
        
        If Geral.Capa.Duplicidade = 1 Then bDuplicidadeCapa = True
        
        'Inicia Transação
        Geral.Banco.BeginTrans
       
        'Finaliza atualização da Capa devido ao encerramento de Capa
        TmrAtualiza.Enabled = False

        With Modulo.qryRemoveCapaRecepcionada
            .rdoParameters(1) = Geral.DataProcessamento
            .rdoParameters(2) = Geral.Capa.Capa
            .rdoParameters(3) = Geral.Capa.AgOrig
            .rdoParameters(4) = Geral.Capa.Num_Malote
            .Execute
            If .rdoParameters(0).Value <> 0 Then
                Geral.Banco.RollbackTrans
                'Retorna/Inicia o controle de tempo para informar que capa em complementação
                TmrAtualiza.Enabled = True
                
                Beep
                MsgBox "Não foi possível atualizar documento referente a Capa.", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        End With

        'Encerra Envelope
        If Not EncerraCapa() Then
            Geral.Banco.RollbackTrans
            'Retorna/Inicia o controle de tempo para informar que capa em complementação
            TmrAtualiza.Enabled = True
            
            Beep
            MsgBox "Não foi possível encerrar o envelope anterior!", vbInformation + vbOKOnly, App.Title
            iSituacao = 1: Exit Function
        End If
        
        'Faz o split de envelope (Inclui Nova Capa com status = 2 e atualiza Docto)
         Geral.Capa.IdEnv_Mal = "M"
         If Not SplitCapa() Then
            Geral.Banco.RollbackTrans
            'Retorna/Inicia o controle de tempo para informar que capa em complementação
            TmrAtualiza.Enabled = True
            
            Beep
            MsgBox "Não foi possível fazer o split de envelope !", vbInformation + vbOKOnly, App.Title
            iSituacao = 1: Exit Function
        End If
        
        'Retorna/Inicia o controle de tempo para informar que capa em complementação
        TmrAtualiza.Enabled = True
        
        'Finaliza transação
        Geral.Banco.CommitTrans
        
        '------------------------------------------
        '---    Atualiza Grid de Navegação      ---
        '------------------------------------------
        'Se Split efetuado com sucesso, atualiza IdCapa de todos documento referentes ao novo Envelope/Malote
        iPosicaoAtualGrid = grdDocumentos.Row
        'Informa no primeiro Docto do novo envelope que Tipo de Docto é (1)Envelope/Malote
        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = 1
        grdDocumentos.Col = iColLeitura: grdDocumentos.Text = Geral.Documento.Leitura
        grdDocumentos.Col = iColStatus: grdDocumentos.Text = 1  'Complementado
        
        'Posiciona na coluna contendo informação de IdCapa
        grdDocumentos.Col = iColdCapa
        Do
            grdDocumentos.Text = Geral.Capa.IdCapa
            If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then Exit Do
            grdDocumentos.Row = grdDocumentos.Row + 1
        Loop
            
        'Elimina documento do outro Envelope/Malote (IdCapa)
        grdDocumentos.Row = 0
        grdDocumentos.Col = iColdCapa
        Do
            If grdDocumentos.Text <> Geral.Capa.IdCapa Then
                grdDocumentos.RemoveItem (grdDocumentos.Row)
                If grdDocumentos.Rows = 0 Then Exit Do
            Else
                If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then Exit Do
                grdDocumentos.Row = grdDocumentos.Row + 1
            End If
        Loop
        
        'Atualiza variaveis de Ambiente Atual
        Modulo.Capa.AgOrig = Geral.Capa.AgOrig
        Modulo.Capa.Capa = Geral.Capa.Capa
        Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
        Modulo.Capa.IdCapa = Geral.Capa.IdCapa
        Modulo.Capa.IdEnv_Mal = Geral.Capa.IdEnv_Mal
        Modulo.Capa.Num_Malote = Geral.Capa.Num_Malote
        Modulo.Capa.Status = Geral.Capa.Status
        
        Modulo.Capa.agefsdtmvan = Geral.Capa.agefsdtmvan
        Modulo.Capa.agefsdtmvat = Geral.Capa.agefsdtmvat
        Modulo.Capa.agefsestado = Geral.Capa.agefsestado
        Modulo.Capa.agefsstmovi = Geral.Capa.agefsstmovi
        bSupervisor = False
        grdDocumentos.Row = 0
        
    End With

    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function


Err_ComplementaMalote:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload Malote
    Geral.Banco.RollbackTrans
   
    Select Case TratamentoErro("1. Não foi possível complementar Malote.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Public Sub cmdFrenteVerso_Click()

On Error GoTo Err_Imagem
    Dim Ret As Long
    With Lead1
        .AutoRepaint = False
    
        'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
        'poi, o canon não gera verso.
        If .Tag = "1" Or (Geral.Documento.Ordem = "1" And .Tag = "0") Then
            'Se verso, mostrar frente
            .Tag = "0"
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Geral.Documento.Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Frente, 0, 0, 1
            End If
        Else
            'Se frente, mostrar verso
            .Tag = "1"
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Geral.Documento.Verso, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Verso, 0, 0, 1
            End If
        End If
       
       ' se imagem for da ls500, mostra mais escura
       If Geral.Documento.Ordem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for da canon, mostra a 50%
       If Geral.Documento.Ordem <> "1" Then
          .PaintZoomFactor = 100
       Else
          .PaintZoomFactor = 50
       End If
        
        .AutoRepaint = True
    End With
    iFlagRotacao = 0
    'Posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
    
    Exit Sub

Err_Imagem:
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Public Sub cmdInverteCor_Click()
    
On Error GoTo Err_Imagem
    
    Lead1.Invert
    
    Exit Sub
    
Err_Imagem:
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Public Sub cmdRotacao_Click()
    
On Error GoTo Err_Imagem
    
    Lead1.FastRotate 90
    
    'Rotaciona imagem e controla o número de graus para poder atualizar o arquivo de imagem
    ' Graus (0)-Zero/360 (1)-90 (2)-180 (3)-270
    If iFlagRotacao = 3 Then
        iFlagRotacao = 0
    Else
        iFlagRotacao = iFlagRotacao + 1
    End If

    Exit Sub
    
Err_Imagem:
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Public Sub cmdZoomMais_Click()
    
On Error GoTo Err_Imagem
    
    If Lead1.PaintZoomFactor <= 400 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
    End If
    
    Exit Sub
    
Err_Imagem:
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Public Sub cmdZoomMenos_Click()
    
On Error GoTo Err_Imagem
    
    If Lead1.PaintZoomFactor >= 20 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor - 10
    End If
    
    Exit Sub
    
Err_Imagem:
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Private Function ComplementaEnvelope(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação da Capa de Envelope
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------
Dim iPosicaoAtualGrid As Integer
sPosicaoErro = "ComplEnvelope"
Dim bDuplicidadeCapa As Boolean

On Error GoTo Err_ComplementaEnvelope
    
    sCapaOuDocumento = "D"

    'Carrega variaveis de Capa para complementar
    Geral.Capa.Capa = IIf(Len(Trim(Geral.Documento.Leitura)) > 8, 0, Val(Geral.Documento.Leitura))
    Geral.Capa.Duplicidade = 0
    Geral.Capa.Num_Malote = 0
    bDuplicidadeCapa = False
    
    Load Envelope
    With Envelope
        
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5400
        .Show vbModal, Me
        
        'Fecha Form
        Unload Envelope
        
        If Not .Alterou Then
            iSituacao = 1: Exit Function
        End If
        
        If Geral.Capa.Duplicidade = 1 Then bDuplicidadeCapa = True
        
        'Inicia Transação
        Geral.Banco.BeginTrans
        
        'Finaliza atualização da Capa devido ao encerramento de Capa
        TmrAtualiza.Enabled = False
        
        With Modulo.qryRemoveCapaRecepcionada
            .rdoParameters(1) = Geral.DataProcessamento
            .rdoParameters(2) = Geral.Capa.Capa
            .rdoParameters(3) = Geral.Capa.AgOrig
            .rdoParameters(4) = Geral.Capa.Num_Malote
            .Execute
            If .rdoParameters(0).Value <> 0 Then
                Geral.Banco.RollbackTrans
                'Retorna/Inicia o controle de tempo para informar que capa em complementação
                TmrAtualiza.Enabled = True
                
                Beep
                MsgBox "Não foi possível atualizar documento referente a Capa.", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        End With
        
        'Encerra Envelope
        If Not EncerraCapa() Then
            Geral.Banco.RollbackTrans
            'Retorna/Inicia o controle de tempo para informar que capa em complementação
            TmrAtualiza.Enabled = True
            
            Beep
            MsgBox "Não foi possível encerrar o envelope anterior!", vbInformation + vbOKOnly, App.Title
            iSituacao = 1: Exit Function
        End If
        
        'Faz o split de envelope (Inclui Nova Capa com status = 2 e atualiza Docto)
        Geral.Capa.IdEnv_Mal = "E"
        If Not SplitCapa() Then
            Geral.Banco.RollbackTrans
            'Retorna/Inicia o controle de tempo para informar que capa em complementação
            TmrAtualiza.Enabled = True
            
            Beep
            MsgBox "Não foi possível fazer o split de envelope !", vbInformation + vbOKOnly, App.Title
            iSituacao = 1: Exit Function
        End If

        'Retorna/Inicia o controle de tempo para informar que capa em complementação
        TmrAtualiza.Enabled = True
        
        
        'Finaliza transação
        Geral.Banco.CommitTrans
        
        '------------------------------------------
        '---    Atualiza Grid de Navegação      ---
        '------------------------------------------
        'Se Split efetuado com sucesso, atualiza IdCapa de todos documento referentes ao novo Envelope/Malote
        iPosicaoAtualGrid = grdDocumentos.Row
        'Informa no primeiro Docto do novo envelope que Tipo de Docto é (1)Envelope/Malote
        grdDocumentos.Col = iColTpDocto: grdDocumentos.Text = 1
        grdDocumentos.Col = iColLeitura: grdDocumentos.Text = Geral.Capa.Capa
        grdDocumentos.Col = iColStatus: grdDocumentos.Text = 1  'Complementado
        
        'Posiciona na coluna contendo informação de IdCapa
        grdDocumentos.Col = iColdCapa
        Do
            grdDocumentos.Text = Geral.Capa.IdCapa
            If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then Exit Do
            grdDocumentos.Row = grdDocumentos.Row + 1
        Loop
            
        'Elimina documento do outro Envelope/Malote (IdCapa)
        grdDocumentos.Row = 0
        grdDocumentos.Col = iColdCapa
        Do
            If grdDocumentos.Text <> Geral.Capa.IdCapa Then
                grdDocumentos.RemoveItem (grdDocumentos.Row)
                If grdDocumentos.Rows = 0 Then Exit Do
            Else
                If grdDocumentos.Row = (grdDocumentos.Rows - 1) Then Exit Do
                grdDocumentos.Row = grdDocumentos.Row + 1
            End If
        Loop
        
        'Atualiza variaveis de Ambiente Atual
        Modulo.Capa.AgOrig = Geral.Capa.AgOrig
        Modulo.Capa.Capa = Geral.Capa.Capa
        Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
        Modulo.Capa.IdCapa = Geral.Capa.IdCapa
        Modulo.Capa.IdEnv_Mal = Geral.Capa.IdEnv_Mal
        Modulo.Capa.Num_Malote = Geral.Capa.Num_Malote
        Modulo.Capa.Status = Geral.Capa.Status
        
        Modulo.Capa.agefsdtmvan = Geral.Capa.agefsdtmvan
        Modulo.Capa.agefsdtmvat = Geral.Capa.agefsdtmvat
        Modulo.Capa.agefsestado = Geral.Capa.agefsestado
        Modulo.Capa.agefsstmovi = Geral.Capa.agefsstmovi
        
        bSupervisor = False
        grdDocumentos.Row = 0
        
    End With

    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaEnvelope:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload Envelope
    Geral.Banco.RollbackTrans
   
    Select Case TratamentoErro("1. Não foi possível complementar Envelope.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaCheque(iSituacao As Integer)
'--------------------------------------------------------------------------------------
'   Complementação de Cheque
'   Parâmetro:  iSituacao   -   (9)-Cheque com valor impresso, dispensa digitação
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'--------------------------------------------------------------------------------------

Dim bExecutou As Boolean

sPosicaoErro = "ComplCheque"

On Error GoTo Err_ComplementaCheque

    'Verifica se complementação automática
    If iSituacao = 9 Then
        sPosicaoErro = "ComplChequeAutom"
        Screen.MousePointer = vbHourglass
        
        iSituacao = 9
        
        On Error GoTo Err_ComplementaChequeAuto
        
        'Inicia Transação
        Geral.Banco.BeginTrans
        
        Call ComplementaChequeAutomatico(iSituacao)
        Screen.MousePointer = vbDefault

        'Verifica se ocorreu erro na execução de SP
        If iSituacao <> 0 Then
            Geral.Banco.RollbackTrans
        End If
        
        If iSituacao = 2 Then GoTo Err_ComplementaChequeAuto
        If iSituacao = 1 Then Exit Function
        
        'Verificar, se houve inversão de imagem atualiza documento
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        
        If Geral.Documento.Status <> "D" Then
            Geral.Documento.Status = "C"    'Complementado (Não apresenta no Docto Anterior)
            
            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
                If Not GeraOcorrenciaDocto Then
                    Geral.Banco.RollbackTrans
                    iSituacao = 1: Exit Function
                End If
            Else
                'Atualiza status documento
                If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1") Then
                    Geral.Banco.RollbackTrans
                    iSituacao = 1: Exit Function
                End If
            End If
        Else
            'Atualiza status documento
            If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "D") Then
                Geral.Banco.RollbackTrans
                iSituacao = 1: Exit Function
            End If
        End If
        
        'Finaliza transação
        Geral.Banco.CommitTrans
       
    Else
        'Multiplica valor por 100 para obter valor decimal
        Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
        Load Cheque
        With Cheque
            .SetParent Complementacao
            .SetPosition ((Screen.Width - .Width) / 2), 5800
            
            .Show vbModal, Me
        
            'Fecha Form
            Unload Cheque
            
            If .Alterou Then
                If Geral.Documento.Status <> "D" Then
                    Geral.Documento.Status = "1"
                    'Verificar, se houve inversão de imagem atualiza documento
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    
                    If Lead1.Tag <> "0" Then
                        bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    Else
                        bExecutou = True
                    End If
                    
                    'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                    If Geral.Capa.Duplicidade <> 0 Then
                        If Not GeraOcorrenciaDocto Then bExecutou = False
                    End If
            
                    'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                    If bExecutou And iFlagRotacao <> 0 Then
                        bExecutou = AlteraRotacao
                    End If
                    
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                Else
                    'Verificar, se houve inversão de imagem atualiza documento
                    
                    If Lead1.Tag = "1" Then
                        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                        bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                        If Not bExecutou Then
                            Beep
                            MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                            iSituacao = 1: Exit Function
                        End If
                    End If
                
                End If
            Else
                iSituacao = 1: Exit Function
            End If
        End With
    End If
    
    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocChequeUBBSacado
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function


Err_ComplementaCheque:
    
    'Finaliza complementação com Erro
    Unload Cheque

Err_ComplementaChequeAuto:
    iSituacao = 2
    Screen.MousePointer = vbDefault
    Geral.Banco.RollbackTrans
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento Cheque.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function AlteraRotacao() As Boolean
    
AlteraRotacao = False

On Error GoTo Err_Imagem

    If iFlagRotacao <> 0 Then
        If Lead1.Tag = "0" Then
            'Carrega a LEAD novamente para não perder a intensidade da imagem
            If Geral.VIPSDLL = eDllProservi Then
              Lead1.Load Geral.DiretorioImagens & Geral.Documento.Frente, 0, 0, 1
            Else
              Lead1.Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Frente, 0, 0, 1
            End If

            'Rotaciona a imagem carregada conforme operação do usuário
            Lead1.FastRotate 90 * iFlagRotacao
            'Salva a imagem rotacionada
            If Geral.VIPSDLL = eDllProservi Then
              Lead1.Save Geral.DiretorioImagens & Geral.Documento.Frente, 10, 8, 80, False
            Else
              Lead1.Save Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Frente, 10, 8, 80, False
            End If
        Else
            'Carrega a LEAD novamente para não perder a intensidade da imagem
            If Geral.VIPSDLL = eDllProservi Then
              Lead1.Load Geral.DiretorioImagens & Geral.Documento.Verso, 0, 0, 1
            Else
              Lead1.Load Geral.DiretorioImagens & Format(Geral.Capa.IdLote, "000000000") & "\" & Geral.Documento.Verso, 0, 0, 1
            End If

            'Rotaciona a imagem carregada conforme operação do usuário
            Lead1.FastRotate 90 * iFlagRotacao
            'Salva a imagem rotacionada
            Lead1.Save Geral.DiretorioImagens & Geral.Documento.Verso, 10, 8, 80, False
        End If
        
        iFlagRotacao = 0

    End If
    
    AlteraRotacao = True
    Exit Function
    
Err_Imagem:
    Beep
    MsgBox "Não foi possível alterar rotação da imagem do documento" & vbCrLf & vbCrLf & _
            "Favor contatar o suporte com as informações abaixo" & vbCrLf & _
            "(" & Geral.DiretorioImagens & Geral.Documento.Frente & ")" & vbCrLf & _
            "(" & Geral.DiretorioImagens & Geral.Documento.Verso & ")" & vbCrLf & _
            "Tag = " & Lead1.Tag, vbExclamation + vbOKOnly, App.Title
    
End Function
Private Function ComplementaADCC(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Autorização de Débito
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplADCC"

On Error GoTo Err_ComplementaADCC

    Load ADCC
    With ADCC
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5800
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload ADCC
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                If Geral.Documento.Status <> "L" Then Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status)
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status, Geral.Documento.Verso, Geral.Documento.Frente)
                End If
                
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                    If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocADCC
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaADCC:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload ADCC
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaArrecConvencional(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Arrecadação Convencional
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplArrecConv"

On Error GoTo Err_ComplementaArrecConvencional

    'Multiplica valor por 100 para obter valor decimal
    Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
    Load ArrecConvencional
    With ArrecConvencional
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5700
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload ArrecConvencional
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocArrecConvencional
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaArrecConvencional:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload ArrecConvencional
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaDeposito(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação da Ficha de Depósito
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplDeposito"

On Error GoTo Err_ComplementaDeposito

    'Multiplica valor por 100 para obter valor decimal
    Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
    Load Deposito
    With Deposito
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5180
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload Deposito
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                If Geral.Documento.Status <> "L" Then Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura

                If Lead1.Tag <> "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status, Geral.Documento.Verso, Geral.Documento.Frente)
                Else
                    bExecutou = True
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                    If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocDepositoCC
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaDeposito:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload Deposito
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaFichaCompensacao(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação da Ficha de Compensação
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplFichaCompensa"

On Error GoTo Err_ComplementaFichaCompensacao

    'Multiplica valor por 100 para obter valor decimal
    Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
    Load FichaCompensacao
    With FichaCompensacao
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5350
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload FichaCompensacao
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                
                If Lead1.Tag <> "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                Else
                    bExecutou = True
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocFichaCompensacao
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaFichaCompensacao:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload FichaCompensacao
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaTitulos(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Títulos
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplTitulos"

On Error GoTo Err_ComplementaTitulos

    'Multiplica valor por 100 para obter valor decimal
    Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
    Load Titulo
    With Titulo
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5800
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload Titulo
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If
                
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocTitulos
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaTitulos:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload Titulo
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function

Private Sub tmrAtualiza_Timer()
  
    TmrAtualiza.Enabled = False
    
    If Geral.Capa.IdCapa <> 0 Then
        sTempo = sTempo + Int(TmrAtualiza.Interval / 1000)

        If sTempo + Int(TmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            'Atualizar o Status da Capa
            Call AtualizaStatusCapa(Geral.Capa.IdCapa, "2")

            sTempo = 0
        End If
    End If
    
    TmrAtualiza.Enabled = True
    
End Sub

Private Function IdentificaTipoDocto(sLeitura As String) As Integer
    Dim ver_cb_arrec As String
    Dim ret_cmc7 As String
    Dim sLeituraAcerto As String, i As Integer
    
    ver_cb_arrec = ""
    
    'Default é zero (Necessita Escolher o Documento)
    IdentificaTipoDocto = 0
    
    If Len(Trim(sLeitura)) = 0 Then Exit Function
    
    'Se Campo Leitura NÂO É totalmente numérico, zera o restante do campo leitura
    If Not IsNumeric(sLeitura) Then
        sLeituraAcerto = ""
        For i = 1 To Len(sLeitura)
            If IsNumeric(Mid(sLeitura, i, 1)) Then
                sLeituraAcerto = sLeituraAcerto + CStr(Mid(sLeitura, i, 1))
            Else
                sLeituraAcerto = sLeituraAcerto + String(Len(sLeitura) - Len(sLeituraAcerto), "0")
                Exit For
            End If
        Next
        'Atualiza Campo Leitura do Grid de Navegação
        grdDocumentos.Col = iColLeitura
        grdDocumentos.Text = sLeituraAcerto
        'Retorna como documento indefnido
        Exit Function
    End If
    
    'Verifica se Documento é uma Capa de  Envelope
    If Len(Trim(sLeitura)) = 8 Then
        IdentificaTipoDocto = 1
        Exit Function
    End If
    'Verifica se Documento é uma Capa de  Envelope
    If Len(Trim(sLeitura)) = 14 Then
        IdentificaTipoDocto = 99
        Exit Function
    End If
        
    If Len(Trim(sLeitura)) = 44 Then

        'Documento de Código de Barras
        If Left(sLeitura, 1) = "8" Then ' É Concessionária
            If VerificarArrecadacaoConvencional(sLeitura) Then
            
                'Arrecadação Convencional
                IdentificaTipoDocto = 27
                
            ElseIf Mid(sLeitura, 3, 1) = "6" Then

                'Valor em Real
                ver_cb_arrec = Mid(sLeitura, 4, 1) + Mid(sLeitura, 1, 3) + _
                               Mid(sLeitura, 5, 40)
                If Not Modulo10Arrecadacao(ver_cb_arrec, 44) Then
                    
                    'Necessita Escolher o Tipo de Documento
                    Exit Function
                Else
                    If Left(sLeitura, 2) = "81" Then
                        If Val(Mid(sLeitura, 20, 2)) = 23 Then
                            IdentificaTipoDocto = 999    'Tributo Imobiliário - Não necessita digitação
                        Else
                            IdentificaTipoDocto = 24     'Valor Indexado
                        End If
                    ElseIf Left(sLeitura, 2) = "86" Then
                        IdentificaTipoDocto = 25     'Trib. Estadual - DPVAT
                    ElseIf Left(sLeitura, 2) = "85" Then
                        'FGTS
                        If Val(Mid(sLeitura, 17, 3)) = 107 Or Val(Mid(sLeitura, 17, 3)) = 108 Or _
                            Val(Mid(sLeitura, 17, 3)) = 111 Or Val(Mid(sLeitura, 17, 3)) = 112 Then
                          IdentificaTipoDocto = 40
                            'TRIBUTOS
                        ElseIf Val(Mid(sLeitura, 17, 3)) <= 27 Then
                          'TRIBUTOS ESTADUAIS
                          IdentificaTipoDocto = 25
                        Else
                          'TRIBUTOS FEDERAIS
                          IdentificaTipoDocto = 26
                        End If
                    
                    Else
                        'Documento não precisa de Digitação
                        IdentificaTipoDocto = 999
                    End If
                End If
            
            ElseIf Mid(sLeitura, 2, 1) = "5" Then
            
                ver_cb_arrec = Mid(sLeitura, 4, 1) + Mid(sLeitura, 1, 3) + _
                               Mid(sLeitura, 5, 40)
                If Not Modulo10Arrecadacao(ver_cb_arrec, 44) Then
                    'Necessita Escolher o Tipo de Documento
                    Exit Function
                Else
                    'FGTS
                    If Val(Mid(sLeitura, 17, 3)) = 107 Or Val(Mid(sLeitura, 17, 3)) = 108 Or _
                        Val(Mid(sLeitura, 17, 3)) = 111 Or Val(Mid(sLeitura, 17, 3)) = 112 Then
                        IdentificaTipoDocto = 40
                    Else
                        'Arrecadação Convencional
                        IdentificaTipoDocto = 27
                    End If
                End If
                
            ElseIf (Mid(sLeitura, 2, 1) = "2") And (Mid(sLeitura, 3, 1) = "4") And (Mid(sLeitura, 16, 4) = "0014") Then
                'CDAE-RJ - Documento não precisa de Digitação
                IdentificaTipoDocto = 999
            
            ElseIf Mid(sLeitura, 2, 2) = "17" And (Val(Mid(sLeitura, 20, 2)) = 43 Or _
                    Val(Mid(sLeitura, 20, 2)) = 52 Or Val(Mid(sLeitura, 20, 2)) = 23) Then
                'Títulos imobiliários - Arrec. Municipal (Documento não precisa de Digitação)
               IdentificaTipoDocto = 24
            Else    'Valor Indexado
                IdentificaTipoDocto = 9
            End If
        Else

            'Ficha de Compensação
            IdentificaTipoDocto = 10
        End If
    
    ElseIf Len(Trim(sLeitura)) >= 30 And _
        Len(Trim(sLeitura)) < 40 Then
        
        If Mid(sLeitura, 9, 3) = "600" And _
            Mid(sLeitura, 18, 1) = "4" And _
            Mid(sLeitura, 12, 6) = Mid(sLeitura, 24, 6) Then
            IdentificaTipoDocto = 99
            Exit Function
        End If

        Select Case Mid(sLeitura, 20, 3)
                            
                Case "999"  'Ficha de Depósito
                    IdentificaTipoDocto = 2
                
                Case "256"  'Ficha de ADCC
                    IdentificaTipoDocto = 4
                    
                Case "592"  'Capa de OCT
                    'Documento não precisa de Digitação
    
                    ' verificar o cmc-7 '
                    ret_cmc7 = Modulo10CMC7(Mid(sLeitura, 1, 30))
                    If ret_cmc7 <> "111" Then
                        IdentificaTipoDocto = 39      'vai para digitação
                    Else
                        IdentificaTipoDocto = 888     'não vai para digitação
                    End If
                
                Case Else
                        
                    'caso o usuario volte um docto de ADCC, o mesmo estava voltando como cheque.
                    If Mid(sLeitura, 1, 30) = "409000000000000000000000000000" Or _
                       Mid(sLeitura, 1, 30) = "230000000000000000000000000000" Then
                        IdentificaTipoDocto = 4
                    Else
                        'Cheque
                        If Mid(sLeitura, 18, 1) = "5" Or Mid(sLeitura, 18, 1) = "8" Or _
                            Mid(sLeitura, 18, 1) = "6" Or Mid(sLeitura, 18, 1) = "9" Then
                            If Left(sLeitura, 3) = "409" Or Left(sLeitura, 3) = "230" Then
                                IdentificaTipoDocto = 5 'Unibanco
                            Else
                                IdentificaTipoDocto = 6 'Outros Bancos
                            End If
                        Else
                            IdentificaTipoDocto = 0 'Nâo definido
                        End If
                    End If
        End Select
    End If

End Function


Private Function DefineTipoDocto(ByVal sCodigoBarras As String) As Integer

  'Verificar se o documento é uma Arrecadação Eletrônica
  If (Mid(sCodigoBarras, 1, 1) = "8") Then
    'Determinar qual o tipo de Arrecadação
    Select Case (Mid(sCodigoBarras, 2, 1))
      Case "1"
        'TRIBUTOS MUNICIPAIS
        DefineTipoDocto = 24

      Case "2"
        'ÁGUA
        DefineTipoDocto = 20

      Case "3"
        'GÁS OU LUZ
        If Mid(sCodigoBarras, 17, 3) = "056" Or Mid(sCodigoBarras, 17, 3) = "057" Then
           DefineTipoDocto = 21        'GÁS
        Else
           DefineTipoDocto = 22        'LUZ
        End If

      Case "4"
        'TELEFONE
        DefineTipoDocto = 23

      Case "5"
        'FGTS
        If Val(Mid(sCodigoBarras, 17, 3)) = 107 Or Val(Mid(sCodigoBarras, 17, 3)) = 108 Or _
            Val(Mid(sCodigoBarras, 17, 3)) = 111 Or Val(Mid(sCodigoBarras, 17, 3)) = 112 Then
          DefineTipoDocto = 40
          'TRIBUTOS
        ElseIf Val(Mid(sCodigoBarras, 17, 3)) <= 27 Then
          'TRIBUTOS ESTADUAIS
          DefineTipoDocto = 25
        Else
          'TRIBUTOS FEDERAIS
          DefineTipoDocto = 26
        End If

      Case "6"
        'DPVAT
        DefineTipoDocto = 25

      Case Else
        'Sem Código
        DefineTipoDocto = 0
    End Select
  Else
    DefineTipoDocto = 0
  End If
  
End Function

Private Function ComplementaConcessionaria(iSituacao As Integer)
'---------------------------------------------------------------------------------------------
'   Complementação de Concessionárias com Valor (Não necessita digitação do Operador)
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Código de Barras com Duplicidade/Docto não reconhecido
'                               (2)-Erro na Execução
'                               (3)-Enviar para Arrecadação Convencional
'                               (4)-Arrecadação com Código de barras inconsistente
'                                   (Zeros no final de cada Bloco) Somente p/ blocos 2,3 e 4
'---------------------------------------------------------------------------------------------

Dim iTipoDocto As Integer
Dim strEncripta   As String

sPosicaoErro = "ComplConcVlr"

On Error GoTo Err_ComplementaConcessionaria
    'Define o Tipo de Documento conforme o Código de Barras
    iTipoDocto = DefineTipoDocto(Geral.Documento.Leitura)
    
    'Se Tipo de Documento não reconhecido, força Arrecadação eletrônica
    If iTipoDocto = 0 Then iSituacao = 1: Exit Function
    
    'Verifica se existe valor informado no C.Barras
    If Val(Mid(Geral.Documento.Leitura, 5, 11)) = 0 Then iSituacao = 1: Exit Function
    
    If Mid(Geral.Documento.Leitura, 20, 3) = "000" Then iSituacao = 4: Exit Function
      
    If Mid(Geral.Documento.Leitura, 31, 3) = "000" Then iSituacao = 4: Exit Function
    
    If Mid(Geral.Documento.Leitura, 42, 3) = "000" Then iSituacao = 4: Exit Function
    
    'Inicia Transação
    Geral.Banco.BeginTrans
    
    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(iTipoDocto, Geral.Documento.Leitura)
    If strEncripta = "" Then
        Geral.Banco.RollbackTrans
        iSituacao = 1: Exit Function
    End If
    
    'Atualizar / Inserir Arrecadação
    With Modulo.qryAtualizaArrecEletronica
        .rdoParameters(1) = Geral.DataProcessamento                         'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto                         'IdDocto
        .rdoParameters(3) = Geral.Documento.Leitura                         'Codigo de Barras
        .rdoParameters(4) = Mid(Geral.Documento.Leitura, 1, 4)              '1. Parte do Codigo de Barras
        .rdoParameters(5) = Mid(Geral.Documento.Leitura, 16, 29)            '2. Parte do Codigo de Barras
        .rdoParameters(6) = Val(Mid(Geral.Documento.Leitura, 5, 11)) / 100  'Valor
        .rdoParameters(7) = iTipoDocto                                      'TipoDocto
        .rdoParameters(8) = strEncripta                             'Autenticacao digital
        .Execute
    
        'Verifica se houve erro de SP
        If .rdoParameters(0).Value = 3 Then
            Geral.Banco.RollbackTrans
            iSituacao = 2: Exit Function
        End If
        
        'Verifica se documento com duplicidade somente do código de referencia (Sem Valor)
        'enviar para arrecadação convencional
        If .rdoParameters(0).Value = 1 Then
            Geral.Banco.RollbackTrans
            iSituacao = 3: Exit Function
        End If
        
        Geral.Documento.ValorTotal = Val(Mid(Geral.Documento.Leitura, 5, 11)) / 100

        'Verifica se documento em duplicidade
        If .rdoParameters(0).Value = 2 Then
            Geral.Documento.Status = "D"
            'Atualiza somente o Tipo de documento para arrecadação em duplicidade
            'devido ao status já alterado para 'D'
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , iTipoDocto) Then
                Geral.Banco.RollbackTrans
                iSituacao = 1: Exit Function
            End If
        Else
            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
                If Not GeraOcorrenciaDocto Then
                    Geral.Banco.RollbackTrans
                    iSituacao = 1: Exit Function
                End If
            Else
                'Atualiza somente o status para (1)-Complementado
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1") Then
                    Geral.Banco.RollbackTrans
                    iSituacao = 1: Exit Function
                End If
            End If
        End If
    End With

    'Finaliza transação
    Geral.Banco.CommitTrans
    
    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocConcessionariaValorReais
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"    'Devolvido (Não apresenta no Docto Anterior)
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "C"    'Complementado (Não apresenta no Docto Anterior)
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaConcessionaria:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Geral.Banco.RollbackTrans
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaChequeAutomatico(iSituacao As Integer)
'--------------------------------------------------------------------------------------
'   Complementação de Cheque Automatico (Sem necessidade de digitação)
'
'   Retorno:    iSituacao -     (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'--------------------------------------------------------------------------------------

Dim bExecutou As Boolean
Dim sLeitura As String
Dim sCmc71 As String, sCmc72 As String, sCmc73 As String
Dim TipoDocto As String, sSql As String, svalor As String
Dim RsCheque As rdoResultset
Dim strEncripta   As String

On Error GoTo Err_ComplementaChequeAutomatico
    
    sLeitura = Geral.Documento.Leitura
  
    'Verificar se o CMC7 é válido
    If Not (Left(sLeitura, 3) = "409" Or Left(sLeitura, 3) = "230") Then
        iSituacao = 1
        Exit Function
    End If
    
    'Verifica se CMC7 está totalmente correto
    If Not TratarCamposCMC7(sLeitura, sCmc71, sCmc72, sCmc73, svalor) Then
        iSituacao = 1
        Exit Function
    End If

    'Definir o Tipo do Documento
    If Left(sLeitura, 3) = "409" Or Left(sLeitura, 3) = "230" Then
      'Cheque Unibanco
      TipoDocto = "5"
    Else
      'Cheque Terceiro
      TipoDocto = "6"
    End If
    
    'Verificar se Documento está duplicado
    sSql = Geral.DataProcessamento & " , '" & sLeitura & "'," & Geral.Documento.IdDocto
    
    Set Modulo.qryGetChequeDuplicado = Geral.Banco.CreateQuery("", "{call GetChequeDuplicado (" & sSql & ")}")
    
    Set RsCheque = Modulo.qryGetChequeDuplicado.OpenResultset(rdOpenStatic, rdConcurReadOnly)

    If Not RsCheque.EOF Then
      'Atualizar : Status para 'D' , Duplicidade para 1 , Ocorrencia = 998
      With Modulo.qryAtualizaDocumentoExcluido
            .rdoParameters(1) = Geral.DataProcessamento   'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto   'IdDocto
            .rdoParameters(3) = "D"                       'Status
            .rdoParameters(4) = 1                         'Duplicidade
            .rdoParameters(5) = 998                       'Ocorrencia
            .Execute
      End With
    
      Geral.Documento.Status = "D"
    
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(TipoDocto, sLeitura)
    If strEncripta = "" Then
        iSituacao = 1
        Exit Function
    End If
    
    'Inserir / Atualizar registro na tabela 'CHEQUE'
    With Modulo.qryAtualizaCheque
      .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
      .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
      .rdoParameters(3) = sLeitura                    'CMC7
      .rdoParameters(4) = Geral.Documento.ValorTotal  'Valor
      .rdoParameters(5) = TipoDocto                   'TipoDocto
      .rdoParameters(6) = strEncripta                 'Autenticacao digital
      .Execute
    End With

    'Finaliza complementação com sucesso
    iSituacao = 0
    RsCheque.Close
    Exit Function


Err_ComplementaChequeAutomatico:
    
    'Finaliza complementação com Erro (Trata erro na tela de ComplementaCheque)
    RsCheque.Close
    iSituacao = 2

End Function
Private Function ComplementaArrecEletronica(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Arrecadação Eletrônica
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplArrecElet"

On Error GoTo Err_ComplementaArrecEletronica

    Load ArrecEletronica
    With ArrecEletronica
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5800
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload ArrecEletronica
        
        If .Alterou Then
            If Geral.Documento.Status <> "D" Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag <> "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                Else
                    bExecutou = True
                End If
                    
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                    If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                'Verificar, se houve inversão de imagem atualiza documento
                
                If Lead1.Tag = "1" Then
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                End If
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocConcessionariaValorReais
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaArrecEletronica:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload ArrecEletronica
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaCapaOCT(iSituacao As Integer)
'--------------------------------------------------------------------------------------
'   Complementação da Capa de OCT
'   Parâmetro:  iSituacao   -   (9)-Capa com CMC7 preenchido, dispensa digitação
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'--------------------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplCapaOCT"

On Error GoTo Err_ComplementaCapaOCT

    
    'Verifica se complementação automática
    If iSituacao = 9 Then
        On Error GoTo Err_ComplementaCapaOCTAuto
        
        Geral.Documento.Status = "1"    'Complementado (Não apresenta no Docto Anterior)
        'Atualiza documento (Força leitura para verificar se existe duplicidade)
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
         
        'Inicia Transação
        Geral.Banco.BeginTrans
       
        If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , etpdocCapaOCT, Geral.Documento.Leitura, "1") Then
            Geral.Banco.RollbackTrans
            iSituacao = 1: Exit Function
        End If
        
        'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
        If Geral.Capa.Duplicidade <> 0 Then
            If Not GeraOcorrenciaDocto Then
                Geral.Banco.RollbackTrans
                iSituacao = 1: Exit Function
            End If
        End If
        
        'Finaliza transação
        Geral.Banco.CommitTrans
        If bDuplicidade Then bSupervisor = True
        
    Else
    
        Load CapaOCT
        With CapaOCT
            .SetParent Complementacao
            .SetPosition ((Screen.Width - .Width) / 2), 5700
            
            .Show vbModal, Me
        
            'Fecha Form
            Unload CapaOCT
            
            If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = True
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If
                
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                    If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
            Else
                iSituacao = 1: Exit Function
            End If
        End With
    End If
    
    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocCapaOCT
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0
    Exit Function


Err_ComplementaCapaOCT:
    
    'Finaliza complementação com Erro
    Unload CapaOCT

Err_ComplementaCapaOCTAuto:
    iSituacao = 2
    Screen.MousePointer = vbDefault
    Geral.Banco.RollbackTrans
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaOCT(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de OCT
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplOCT"

On Error GoTo Err_ComplementaOCT

    Load OCT
    With OCT
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5200
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload OCT
        
        If .Alterou Then
            If Geral.Documento.Status <> "L" Then Geral.Documento.Status = "1"
            'Verificar, se houve inversão de imagem atualiza documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Lead1.Tag = "0" Then
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status)
            Else
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status, Geral.Documento.Verso, Geral.Documento.Frente)
            End If
            
            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
               If Not GeraOcorrenciaDocto Then bExecutou = False
            End If
            
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocOCT
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaOCT:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload OCT
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Sub IniciarComplementacao()

Static bEntrou As Boolean
'Se houve novo evento do activate, não entra novamente neste processo
If bEntrou Then Exit Sub

bEntrou = True

Dim lContadorAnterior As Long
Dim PauseTime, start

    bFim = False
    lcontador = (Geral.Atualizacao * 1000)
    'Posiciona frame de Aguardo no centro da tela
    picAguardo.Top = (Screen.Height - picAguardo.Height) / 2
    picAguardo.Left = (Screen.Width - picAguardo.Width) / 2
    
    Do While True

        If lcontador = (Geral.Atualizacao * 1000) Then
            
            'Vai para procedimento de complementação
            TmrAtualiza.Enabled = False
            picAguardo.Visible = False
            
            Call Complementar
            
            If bFim Then Exit Do
            
            pgbAguardo.Min = 0
            pgbAguardo.Max = lcontador
            lcontador = 0: lContadorAnterior = 0
            picAguardo.Visible = True
            cmdFechar.SetFocus
            
            ''''''''''''''''''''''''''''''''''''''''''''
            'Grava o log de Inicio de Aguarda Documento'
            ''''''''''''''''''''''''''''''''''''''''''''
            GravaLog 0, 0, 250
            
        End If
        
        PauseTime = 1   'duração
        start = Timer   'iniciar tempo
        Do While Timer < start + PauseTime
            DoEvents
            If bFim Then Exit Do
        Loop

        If bFim Then Exit Do

        lcontador = lcontador + 1000
            
        If picAguardo.Visible Then
            If lcontador <> lContadorAnterior Then
                pgbAguardo.Value = lcontador
                lContadorAnterior = lcontador
            End If
        End If

    Loop

    'Encerra complementção
    bEntrou = False
    Unload Me

End Sub
Private Function ComplemementaArrecValorIndexado(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Arrecadação com Valor Indexado
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplArrecVlrIndex"

On Error GoTo Err_ComplemementaArrecValorIndexado

    'Multiplica valor por 100 para obter valor decimal
    Geral.Documento.ValorTotal = (Geral.Documento.ValorTotal * 100)
    
    Load ArrecValorIndexado
    With ArrecValorIndexado
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5700
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload ArrecValorIndexado
        
        If .Alterou Then

            'Verificar, se houve inversão de imagem atualiza documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            
            If Lead1.Tag <> "0" Then
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
            Else
                bExecutou = True
            End If
            
            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
                If Not GeraOcorrenciaDocto Then bExecutou = False
            End If
            
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With
    
    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocConcessionariaValorIndexado
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    grdDocumentos.Col = iColValor:      grdDocumentos.Text = Geral.Documento.ValorTotal
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplemementaArrecValorIndexado:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload ArrecValorIndexado
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select


End Function
Private Function ComplementaCartaoAvulso(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Cartão Avulso
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplCartão"

On Error GoTo Err_ComplementaCartaoAvulso

    Load CartaoAvulso
    With CartaoAvulso
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5400
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload CartaoAvulso
        
        If .Alterou Then
            Geral.Documento.Status = "1"
            'Verificar, se houve inversão de imagem atualiza documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Lead1.Tag = "0" Then
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
            Else
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
            End If

            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
               If Not GeraOcorrenciaDocto Then bExecutou = False
            End If
            
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocCartaoAvulso
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaCartaoAvulso:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload CartaoAvulso
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select


End Function
Private Function ComplementaGareICMS(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de GARE/ICMS
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplGareIcms"

On Error GoTo Err_ComplementaGareICMS

    Load GareICMS
    With GareICMS
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5200
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload GareICMS
        
        If .Alterou Then
            Geral.Documento.Status = "1"
            'Verificar, se houve inversão de imagem atualiza documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
             If Lead1.Tag = "0" Then
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
            Else
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
            End If

            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
               If Not GeraOcorrenciaDocto Then bExecutou = False
            End If
            
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocGare
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaGareICMS:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload GareICMS
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Function ComplementaCobrancaRegistrada(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de CobrancaRegistrada
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplCobrReg"

On Error GoTo Err_ComplementaCobrancaRegistrada

    Load CobrancaRegistrada
    With CobrancaRegistrada
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5600
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload CobrancaRegistrada
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If
                
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocCobRegistrada
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaCobrancaRegistrada:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload CobrancaRegistrada
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function

Private Function ComplementaCobrancaEspecial(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de CobrancaEspecial
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplCobrEsp"

On Error GoTo Err_ComplementaCobrancaEspecial

    Load CobrancaEspecial
    With CobrancaEspecial
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5350
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload CobrancaEspecial
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If
                
                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocCobEspecial
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaCobrancaEspecial:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload CobrancaEspecial
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Function ComplementaDarm(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de DARM
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplDARM"

On Error GoTo Err_ComplementaDarm

    Load DARM
    With DARM
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5800
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload DARM
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocDarm
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaDarm:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload DARM
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Function ComplementaDarfPreto(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Darf Preto
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplDarfPreto"

On Error GoTo Err_ComplementaDarfPreto

    Load DARFPreto
    With DARFPreto
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5300
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload DARFPreto
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocDarfPreto
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaDarfPreto:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload DARFPreto
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Function ComplementaDarfSimples(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Darf Simples
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplDarfSimples"

On Error GoTo Err_ComplementaDarfSimples

    Load DARFSimples
    With DARFSimples
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5500
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload DARFSimples
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Lead1.Tag = "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                Else
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocDarfSimples
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaDarfSimples:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload DARFSimples
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Function ComplementaGPS(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de GPS
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplGPS"

On Error GoTo Err_ComplementaGPS

    Load GPS
    With GPS
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5400
        'Variável de ambiente onde informa que todos campos deverão ser atualizados
        .AlteraValor = False
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload GPS
        
        If .Alterou Then
                Geral.Documento.Status = "1"
                'Verificar, se houve inversão de imagem atualiza documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                
                If Lead1.Tag <> "0" Then
                    bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                Else
                    bExecutou = True
                End If

                'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                If Geral.Capa.Duplicidade <> 0 Then
                   If Not GeraOcorrenciaDocto Then bExecutou = False
                End If
                
                'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                If bExecutou And iFlagRotacao <> 0 Then
                    bExecutou = AlteraRotacao
                End If
                
                If Not bExecutou Then
                    Beep
                    MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                    iSituacao = 1: Exit Function
                End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocGPS
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaGPS:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload GPS
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function

Private Sub LogAlteraTipoDocto(ByVal iTipoDocumentoAtual As Integer)

'Verifica se Atual documento difere do Tipo de Documento anterior
If Geral.Documento.TipoDocto <> 0 Then
    
    'Verifica se documento é Arrecadação com Valor Indexado
    If InStr("09*24*25*26", Format(Geral.Documento.TipoDocto, "00")) > 0 And _
       InStr("09*24*25*26", Format(iTipoDocumentoAtual, "00")) > 0 Then Exit Sub
    
    'Verifica se documento é Arrecadação Eletronica
    If InStr("20*21*22*23", Format(Geral.Documento.TipoDocto, "00")) > 0 And _
       InStr("20*21*22*23", Format(iTipoDocumentoAtual, "00")) > 0 Then Exit Sub
    
    'Verifica se documento é Ficha de compensação
    If InStr("10*28*30*31", Format(Geral.Documento.TipoDocto, "00")) > 0 And _
       InStr("10*28*30*31", Format(iTipoDocumentoAtual, "00")) > 0 Then Exit Sub
    
    If Geral.Documento.TipoDocto <> iTipoDocumentoAtual Then
        'Grava Log de ocorrência
        Call GravaLog(Geral.Capa.IdCapa, Geral.Documento.IdDocto, LOG.TipoDoctoAlterado)
    End If
End If

End Sub
Private Function ComplementaCapaAutomatico(sTipoDocto As String) As Boolean
'Atualiza Tabela de Agencia e Verifica se Capa em complementação automática está com Duplicidade

ComplementaCapaAutomatico = False

On Error GoTo Err_ComplementaCapaAutomatico

    'Inicia transação
    Geral.Banco.BeginTrans
    
    If sTipoDocto = "M" Then '(( Malote ))
        sPosicaoErro = "AtuDuplCapa"
        With Modulo.qryChecarEnvelope
            .rdoParameters(1) = Geral.DataProcessamento
            .rdoParameters(2) = Null    ' Para malote não há necessidade de verificar por agência
            .rdoParameters(3) = Val(Geral.Documento.Leitura)
            .rdoParameters(5) = Geral.Documento.IdCapa
            .Execute
            
            If .rdoParameters(0).Value <> 0 Then
                Geral.Banco.RollbackTrans
                GoTo Exit_ComplementaCapaAutomatico
            End If

            'Se existe Capa com duplicidade, solicita recadastramento e envia para supervisor
            If .rdoParameters("@Registros") > 0 Then    'Verifica se existe mais de um pois o atual tambem faz parte da checagem
                Geral.Capa.Duplicidade = 1
            Else
                Geral.Capa.Duplicidade = 0
            End If
        End With
        
        If Geral.Capa.Duplicidade = 1 Then
            'Atualiza Status e Duplicidade da tabela Capa
            With Modulo.qryAtualizaDuplicidadeCapa
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Geral.Capa.IdCapa
                .rdoParameters(3) = Geral.Capa.Status
                .rdoParameters(4) = Geral.Capa.Duplicidade  'Identificador de malote em duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value = 1 Then
                    Geral.Banco.RollbackTrans
                    GoTo Exit_ComplementaCapaAutomatico
                End If
            End With
        End If
        
    Else    '(( Envelope ))
        
        sPosicaoErro = "AtuDuplCapa"
        With Modulo.qryChecarEnvelope
            .rdoParameters(1) = Geral.DataProcessamento
            .rdoParameters(2) = Geral.Capa.AgOrig
            .rdoParameters(3) = Val(Geral.Documento.Leitura)
            .rdoParameters(5) = Geral.Documento.IdCapa
            .Execute
            
            If .rdoParameters(0).Value <> 0 Then
                Geral.Banco.RollbackTrans
                GoTo Exit_ComplementaCapaAutomatico
            End If

            'Se existe Capa com duplicidade, solicita recadastramento e envia para supervisor
            If .rdoParameters("@Registros") > 0 Then    'Verifica se existe mais de um pois o atual tambem faz parte da checagem
                Geral.Capa.Duplicidade = 1
            Else
                Geral.Capa.Duplicidade = 0
            End If
        End With
        
        If Geral.Capa.Duplicidade = 1 Then
            'Atualiza Status e Duplicidade da tabela Capa
            With Modulo.qryAtualizaDuplicidadeCapa
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Geral.Capa.IdCapa
                .rdoParameters(3) = Geral.Capa.Status
                .rdoParameters(4) = Geral.Capa.Duplicidade  'Identificador de malote em duplicidade
                .Execute
                
                'Verifica se ocorreu erro na atualização
                If .rdoParameters(0).Value = 1 Then
                    Geral.Banco.RollbackTrans
                    GoTo Exit_ComplementaCapaAutomatico
                End If
            End With
        End If
    End If

    'Finaliza transação
    Geral.Banco.CommitTrans

    ComplementaCapaAutomatico = True
    Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
    
Exit_ComplementaCapaAutomatico:
    Exit Function
    
Err_ComplementaCapaAutomatico:
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível complementar o documento em Modo Automático.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    GoTo Exit_ComplementaCapaAutomatico

End Function
Private Function SplitAnterior() As Integer

'Retorno :  (0) - Split Ok
'           (1) - Erro de leitura de documentos (Não houve alteração de Dados)
'           (2) - Não Foi possível fazer o Split (Houve alteração de Dados)
'
'           (9) - Erro de execução

Dim IdCapa_Anterior As Long, IdLote_Anterior As Long, strDirImgAnterior As String
Dim IdCapa_Atual As Long, IdLote_Atual As Long, strDirImgAtual As String
Dim strImagemFrenteAtual As String, strImagemFrenteAnterior As String
Dim strImagemVersoAtual As String, strImagemVersoAnterior As String

Dim IdDocto_Inicial As Long
Dim bDuplicidade As Boolean
Dim RstDoctoAnterior As rdoResultset
Dim RstCapaAnterior As rdoResultset
Dim sLeitura As String
Dim lIdDocto As Long, iRow As Integer

On Error GoTo Err_SplitAnterior

    SplitAnterior = 1

    '---------------------------------------------------------------------
    '           Obtem Número do primeiro documento da Capa Atual
    '---------------------------------------------------------------------
    Modulo.qryGetPrimeiroDocumentoCapa.rdoParameters(1) = Geral.DataProcessamento
    Modulo.qryGetPrimeiroDocumentoCapa.rdoParameters(2) = Geral.Documento.IdCapa
    Modulo.qryGetPrimeiroDocumentoCapa.Execute
    If Modulo.qryGetPrimeiroDocumentoCapa.rdoParameters(0).Value = 1 Then
        Beep
        MsgBox "Não foi possível ler informações da Capa atual.", vbExclamation + vbOKOnly, App.Title
        GoTo Exit_SplitAnterior
    End If
    
    IdDocto_Inicial = Modulo.qryGetPrimeiroDocumentoCapa.rdoParameters(3).Value

    '-------------------------------------------------
    '           Obtem Número da capa anterior
    '-------------------------------------------------
    Modulo.qryGetIdDocto.rdoParameters(0) = Geral.DataProcessamento
    Modulo.qryGetIdDocto.rdoParameters(1) = (IdDocto_Inicial - 1)
    Set RstDoctoAnterior = Modulo.qryGetIdDocto.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    If RstDoctoAnterior.EOF Then
        Beep
        MsgBox "Não existe Capa anterior para esta data de movimento.", vbExclamation + vbOKOnly, App.Title
        GoTo Exit_SplitAnterior
    End If
    
    'Identificadores da Ultima capa complementada
    IdCapa_Anterior = RstDoctoAnterior!IdCapa
    IdLote_Anterior = RstDoctoAnterior!IdLote
    
    'Guarda o número do IDCapa Atual para controle no GravaLOG
    IdCapa_Atual = Geral.Capa.IdCapa
    IdLote_Atual = Geral.Capa.IdLote
    
    '-------------------------------------------------
    '   Obtem dados complementares da capa anterior
    '-------------------------------------------------
    Modulo.qryDadosCapa.rdoParameters(1) = Geral.DataProcessamento
    Modulo.qryDadosCapa.rdoParameters(2) = IdCapa_Anterior
    Set RstCapaAnterior = Modulo.qryDadosCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    'Verifica se ocorreu erro
    If Modulo.qryDadosCapa.rdoParameters(0).Value <> 0 Then
        Beep
        MsgBox "Não foi possível ler informações da Capa anterior.", vbExclamation + vbOKOnly, App.Title
        GoTo Exit_SplitAnterior
    End If
    
    'Verifica Status da Capa Anterior
    Select Case RstCapaAnterior!Status

        Case "2"    'Capa em Complementação
            MsgBox "Capa Anterior em Complementação, não será permitido enviar documento para capa anterior.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "B"    'Capa em Recaptura
            MsgBox "Capa Anterior em Recaptura, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "E"    'Capa Expedida
            MsgBox "Capa Anterior já expedida, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "F"    'Capa devolvida pelo caixa Robo
            MsgBox "Capa devolvida pelo caixa robô, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "G"    'Capa em Prova Zero
            MsgBox "Capa Anterior em Prova Zero, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "H"    'Capa em Ilegiveis
            MsgBox "Capa Anterior em Ilegíveis, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "I"    'Capa em Alcada
            MsgBox "Capa Anterior em Alçada, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "J"    'Capa em Vinculo Manual
            MsgBox "Capa Anterior em Vínculo Manual, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "K"    'Capa em Expedicao
            MsgBox "Capa Anterior em Expedição, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "M"    'Capa em Confirmação de Ag/Conta
            MsgBox "Capa em confirmação de Ag/Conta, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "N"    'Capa para CSP
            MsgBox "Capa para C.S.P., não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "O"    'Capa em Troca de Ordem
            MsgBox "Capa Anterior em Troca de Ordem, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "P"    'Capa Devolvida pela Preparação
            MsgBox "Capa devolvida pela preparação, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "Q"    'Capa em CSP
            MsgBox "Capa em C.S.P., não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "R"    'Capa para Transmissão
            MsgBox "Capa Anterior para Transmissão, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "S"    'Capa em Transmissão
            MsgBox "Capa Anterior em Transmissão, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "T"    'Capa Transmitida
            MsgBox "Capa Anterior Transmitida, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "U"    'Capa em Confirmação
            MsgBox "Capa Anterior em Confirmação AG./Conta, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "V"    'Capa em Verificação
            MsgBox "Capa Anterior em Verificação, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "X"    'Capa já enviada à ocorrência para Ubb
            MsgBox "Capa Anterior devolvida com ocorrência, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "Y"    'Capa para correção AG/CC
            MsgBox "Capa Anterior para correção de AG./Conta, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "W"    'Capa em Estorno
            MsgBox "Capa em Estorno, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
        Case "Z"    'Capa em Correção de Agência e Conta
            MsgBox "Capa em correção de Ag/Conta, não será permitido complementação.", vbInformation + vbOKOnly, App.Title
            GoTo Exit_SplitAnterior
    
    End Select
    
    'Obtem número do Iddocto referente ao documento contendo dados de capa
    grdDocumentos.Col = iColIdDocto
    lIdDocto = grdDocumentos.Text
    
    'Finaliza controle de tempo
    sTempo = 0
    TmrAtualiza.Enabled = False
    
    '--------------------------------------
    '       Split para Capa Anterior
    '--------------------------------------
    Modulo.qrySplitCapaAnterior.rdoParameters(1) = Geral.DataProcessamento
    Modulo.qrySplitCapaAnterior.rdoParameters(2) = IdCapa_Anterior
    Modulo.qrySplitCapaAnterior.rdoParameters(3) = Geral.Capa.IdCapa
    Modulo.qrySplitCapaAnterior.Execute
    
    'Verifica se ocorreu erro
    If Modulo.qrySplitCapaAnterior.rdoParameters(0) <> 0 Then
        'Reinicia controle de tempo
        TmrAtualiza.Enabled = True
        
        Beep
        MsgBox "Não foi possível fazer o split para Capa Anterior !", vbInformation + vbOKOnly, App.Title
        GoTo Exit_SplitAnterior
    End If
    
    SplitAnterior = 2
    
    'Passa Atualização de Capa em Complementação (Timer) para Capa Anterior
    Geral.Capa.IdCapa = IdCapa_Anterior
    
    'Inicia controle de tempo
    sTempo = 0
    TmrAtualiza.Enabled = True
    
    'Grava Log de ocorrência
    Call GravaLog(IdCapa_Atual, 0, LOG.SplitAnteriorInicial)
    Call GravaLog(IdCapa_Anterior, 0, LOG.SplitAnteriorFinal)
    
    '--------------------------------------------
    '   Carrega documentos da Capa Anterior
    '--------------------------------------------
    Set RstDoctoAnterior = Nothing
    Modulo.qryGetTodosDocumentosCapa(1) = Geral.DataProcessamento
    Modulo.qryGetTodosDocumentosCapa(2) = IdCapa_Anterior
    Set RstDoctoAnterior = Modulo.qryGetTodosDocumentosCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    'Verifica se ocorreu erro
    If Modulo.qryGetTodosDocumentosCapa.rdoParameters(0) <> 0 Then
        Beep
        MsgBox "Não foi possível ler Documentos da Capa Anterior.", vbExclamation + vbOKOnly, App.Title
        GoTo Exit_SplitAnterior
    End If
            
    'Limpa grid contendo dados de documento
    grdDocumentos.Rows = 0
            
    'Carrega Grid com todos documentos da Capa Anterior
    strDirImgAtual = Geral.DiretorioImagens & Right("000000000" & CStr(IdLote_Atual), 9) & "\"
    strDirImgAnterior = Geral.DiretorioImagens & Right("000000000" & CStr(IdLote_Anterior), 9) & "\"
    
    Do Until RstDoctoAnterior.EOF
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Copiar as imagens da capa anterior (Lote atual) para capa anterior (Lote Anterior)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If RstDoctoAnterior!TipoDocto = 0 And RstDoctoAnterior!Status = "0" And RstDoctoAnterior!Valor = 0 Then
           strImagemFrenteAtual = strDirImgAtual & RstDoctoAnterior!Frente
           strImagemVersoAtual = strDirImgAtual & RstDoctoAnterior!Verso
           
           strImagemFrenteAnterior = strDirImgAnterior & RstDoctoAnterior!Frente
           strImagemVersoAnterior = strDirImgAnterior & RstDoctoAnterior!Verso
           
           FileCopy strImagemFrenteAtual, strImagemFrenteAnterior
           FileCopy strImagemVersoAtual, strImagemVersoAnterior
        
        End If
        
        grdDocumentos.Rows = grdDocumentos.Rows + 1
        grdDocumentos.Row = (grdDocumentos.Rows - 1)
        grdDocumentos.Col = iColIdDocto:    grdDocumentos.Text = RstDoctoAnterior!IdDocto
        grdDocumentos.Col = iColdCapa:      grdDocumentos.Text = RstDoctoAnterior!IdCapa
        grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = RstDoctoAnterior!TipoDocto
        grdDocumentos.Col = iColFrente:     grdDocumentos.Text = RstDoctoAnterior!Frente
        grdDocumentos.Col = iColVerso:      grdDocumentos.Text = RstDoctoAnterior!Verso
        grdDocumentos.Col = iColValor:      grdDocumentos.Text = RstDoctoAnterior!Valor
        grdDocumentos.Col = iColOrdem:      grdDocumentos.Text = RstDoctoAnterior!Ordem
        
        grdDocumentos.Col = iColStatus
        
        'Se Documentos é da Capa Anterior, não será possível complementá-los
        If RstDoctoAnterior!IdDocto < lIdDocto Then
            If RstDoctoAnterior!Status = "0" Then
                grdDocumentos.Text = "D"
            Else
                grdDocumentos.Text = RstDoctoAnterior!Status
            End If
        Else
            If RstDoctoAnterior!IdDocto = lIdDocto Then iRow = grdDocumentos.Row
            grdDocumentos.Text = RstDoctoAnterior!Status
        End If
        
        grdDocumentos.Col = iColLeitura:    sLeitura = Trim(RstDoctoAnterior!Leitura)
        grdDocumentos.Text = sLeitura
        
        'Verifica se TODO Campo leitura = "0"
        If Len(sLeitura) > 0 Then
            If sLeitura = String(Len(sLeitura), "0") Then
                grdDocumentos.Text = ""
            End If
        End If
        RstDoctoAnterior.MoveNext
    Loop
    
    'Vai para o registro do Documento Atual
    grdDocumentos.Row = iRow
    
    'Carrega Parte de variáveis globais do Documento
    Call CarregaDocumento
            
    Geral.Capa.IdEnv_Mal = RstCapaAnterior!IdEnv_Mal
    Geral.Capa.Capa = RstCapaAnterior!Capa
    Geral.Capa.Num_Malote = RstCapaAnterior!Num_Malote
    Geral.Capa.AgOrig = RstCapaAnterior!AgOrig
    Geral.Capa.Status = RstCapaAnterior!Status
    Geral.Capa.Duplicidade = RstCapaAnterior!Duplicidade
    
    Geral.Capa.IdLote = RstCapaAnterior!IdLote
    Geral.Capa.IdCapa = IdCapa_Anterior
    
    'Carrega variaveis com situação atual de capa
    Modulo.Capa.IdEnv_Mal = Geral.Capa.IdEnv_Mal
    Modulo.Capa.Capa = Geral.Capa.Capa
    Modulo.Capa.Num_Malote = Geral.Capa.Num_Malote
    Modulo.Capa.AgOrig = Geral.Capa.AgOrig
    Modulo.Capa.Status = Geral.Capa.Status
    Modulo.Capa.Duplicidade = Geral.Capa.Duplicidade
    
    Modulo.Capa.IdLote = Geral.Capa.IdLote
    Modulo.Capa.IdCapa = Geral.Capa.IdCapa
    
    'Apresenta Cabeçalho do Form
    Call Cabecalho(False, grdDocumentos.Row + 1, grdDocumentos.Rows)
        
    SplitAnterior = 0
    
    Exit Function
 

Exit_SplitAnterior:
    If Not (RstDoctoAnterior Is Nothing) Then RstDoctoAnterior.Close
    If Not (RstCapaAnterior Is Nothing) Then RstCapaAnterior.Close
    Exit Function

Err_SplitAnterior:
    SplitAnterior = 9

    Select Case TratamentoErro("Não foi possível fazer o Split da capa anterior.", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    GoTo Exit_SplitAnterior

End Function
Private Function GeraOcorrenciaDocto() As Boolean

On Error GoTo Err_GeraOcorrenciaDocto

    GeraOcorrenciaDocto = False

    With Modulo.qryGeraOcorrenciaDocumento
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto 'Número do IdDocto
        .Execute
        
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value <> 0 Then Exit Function
        
    End With
    
    GeraOcorrenciaDocto = True
    
    Exit Function

Err_GeraOcorrenciaDocto:
    
    Select Case TratamentoErro("Não foi possível atualizar o documento!", Err, rdoErrors, False)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Function
Private Function ComplementaFGTS(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de FGTS
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'                               (3)-Complementacao Automatica Ok
'                               (4)-Documento inválido
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplFGTS"

On Error GoTo Err_ComplementaFGTS

    Dim bCodBarOK       As Boolean
    Dim lDataValidade   As Long
    Dim sCodigo1        As String
    Dim sCodigo2        As String
    Dim sCompetencia    As String
    Dim lCodRec         As Long
    Dim sTipo           As String * 1
    Dim sCNPJ_CEI       As String
    Dim svalor          As String
    Dim strEncripta     As String
    
    bCodBarOK = False
    
    If Len(Trim(Geral.Documento.Leitura)) = 44 Then
        If Modulo10Arrecadacao(Mid(Geral.Documento.Leitura, 4, 1) & _
                               Mid(Geral.Documento.Leitura, 1, 3) & _
                               Mid(Geral.Documento.Leitura, 5, 40), 44) Then
            ' Obtem a data de validade
            lDataValidade = Val(Mid(Geral.Documento.Leitura, 20, 6))
            ' Ano so tem dois digitos
            If Val(Mid(Geral.Documento.Leitura, 20, 2)) >= 50 Then
                lDataValidade = 19000000 + lDataValidade
            Else
                lDataValidade = 20000000 + lDataValidade
            End If
            
            If lDataValidade >= Geral.DataProcessamento Then
                bCodBarOK = True
            End If
            
        End If
    End If
    
    ' Obtem CNPJ ou CEI, se CNPJ calcular os DV's
    If (Mid(Geral.Documento.Leitura, 32, 1) = "0" Or Mid(Geral.Documento.Leitura, 32, 1) = "3") Then
        iSituacao = 4: Exit Function
    End If
    
    If Mid(Geral.Documento.Leitura, 17, 3) <> "107" Or Not bCodBarOK Then
       
        Load frmFGTS
        With frmFGTS
            .SetParent Complementacao
            .SetPosition ((Screen.Width - .Width) / 2), 5400
            'Variável de ambiente onde informa que todos campos deverão ser atualizados
            .AlteraValor = False
            
            .Show vbModal, Me
        
            'Fecha Form
            Unload frmFGTS
            
            If .Alterou Then
                    Geral.Documento.Status = "1"
                    'Verificar, se houve inversão de imagem atualiza documento
                    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                    If Lead1.Tag = "0" Then
                        bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1")
                    Else
                        bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1", Geral.Documento.Verso, Geral.Documento.Frente)
                    End If
                    
                    'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
                    If Geral.Capa.Duplicidade <> 0 Then
                       If Not GeraOcorrenciaDocto Then bExecutou = False
                    End If
                    
                    'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
                    If bExecutou And iFlagRotacao <> 0 Then
                        bExecutou = AlteraRotacao
                    End If
                    
                    If Not bExecutou Then
                        Beep
                        MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                        iSituacao = 1: Exit Function
                    End If
                    'complementacao manual com sucesso
                    iSituacao = 0
            Else
                iSituacao = 1: Exit Function
            End If
        End With
    Else
        ' FGTS cod 107 pode ser complementado automaticamente
        
        ' Calcular a Competencia
        ' pega o codigo da competencia no codigo de barras
        sCodigo1 = Mid(Geral.Documento.Leitura, 26, 3)
        ' conforme a regra
        sCodigo2 = DateDiff("m", "31/12/1966", Format(Date, "dd/mm/yyyy"))
        ' formata como mm/yyyy
        sCompetencia = Format(DateSerial(Year(Date), Month(Date) - (Val(sCodigo2) - Val(sCodigo1)), 1), "yyyymm")
        
        ' Obtem o Codigo de Recolhimento
        ' pega o codigo de recolhimento no codigo de barras
        lCodRec = Val(Mid(Geral.Documento.Leitura, 29, 3))
        
        ' Obtem CNPJ ou CEI, se CNPJ calcular os DV's
        sTipo = Mid(Geral.Documento.Leitura, 32, 1)
        ' Se Tipo = 1 entao e CNPJ
        If sTipo = "1" Then
            ' carrega o form so para usar a funcao CalculaCGC, mas ele nao sera exibido
            Load frmFGTS
            sCNPJ_CEI = Format(Right(Geral.Documento.Leitura, 12) & "00", String(15, "0"))
            sCNPJ_CEI = frmFGTS.CalculaCGC(sCNPJ_CEI)
            sCNPJ_CEI = Mid(sCNPJ_CEI, 2, 14)
            Unload frmFGTS
        Else
            sCNPJ_CEI = Right(Geral.Documento.Leitura, 12)
        End If
        
        ' Obtem valor do deposito
        svalor = Mid(Geral.Documento.Leitura, 5, 11)
    
        On Error GoTo ERRO_GravarDocumento:
        
        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(Geral.Documento.TipoDocto, sCNPJ_CEI)
        If strEncripta = "" Then
            iSituacao = 1
            Exit Function
        End If
        
        Set Modulo.qryAtualizaFGTS = Geral.Banco.CreateQuery("", "{? = call AtualizaFGTS (?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
        
        With Modulo.qryAtualizaFGTS
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.DataProcessamento                 'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto                 'IdDocto
            .rdoParameters(3) = Geral.Documento.Leitura                 'Codigo de Barras
            .rdoParameters(4) = lCodRec                                 'Codigo de recolhimento
            .rdoParameters(5) = sCNPJ_CEI                               'CGC da Empresa
            .rdoParameters(6) = sCompetencia                            'Competencia
            .rdoParameters(7) = lDataValidade                           'Data de Validade
            .rdoParameters(8) = ""                                      'CGC do Tomador
            .rdoParameters(9) = Val(svalor) / 100                       'Valor Deposito
            .rdoParameters(10) = 0                                      'Valor JAM
            .rdoParameters(11) = 0                                      'Valor da Multa
            .rdoParameters(12) = Val(svalor) / 100                      'Valor total
            .rdoParameters(13) = 40                                     'Tipo de documento
            .rdoParameters(14) = strEncripta                            'Autenticacao digital
            .Execute
        End With
        
        If Modulo.qryAtualizaFGTS.rdoParameters(0) = 1 Then
            GoTo ERRO_GravarDocumento:
        End If
        If Modulo.qryAtualizaFGTS.rdoParameters(0) = 2 Then
            'Documento Duplicado
            Geral.Documento.Status = "D"
            Geral.Capa.Duplicidade = "1"
        End If
        Modulo.qryAtualizaFGTS.Close

        
        '''''''''''''''''''''''''''''
        'Atualizar o Controle Global
        '''''''''''''''''''''''''''''
        Geral.Documento.ValorTotal = Val(svalor) / 100
        Geral.Documento.TipoDocto = 40
        Geral.Documento.Status = "1"
        
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , "1") Then
            iSituacao = 1
            Exit Function
        End If
        
        'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
        If Geral.Capa.Duplicidade <> 0 Then
           If Not GeraOcorrenciaDocto Then
              iSituacao = 1
              Exit Function
          End If
        End If
        
        'complementacao automatica com sucesso
        iSituacao = 3
        
    End If
    
    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocFGTS
    grdDocumentos.Col = iColLeitura:    grdDocumentos.Text = Geral.Documento.Leitura
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    Exit Function

ERRO_GravarDocumento:
    iSituacao = 2
    Call TratamentoErro("Erro ao Atualizar o documento.", Err, rdoErrors, False)
    Exit Function

Err_ComplementaFGTS:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload frmFGTS
   
    Call TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)

End Function

Private Function ComplementaLancamentoInterno(iSituacao As Integer)
'----------------------------------------------------------------------------
'   Complementação de Lancamento Interno
'
'   Retorno:    iSituacao   -   (0)-Complementação Ok
'                               (1)-Cancelou Complementação
'                               (2)-Erro na Complementação
'----------------------------------------------------------------------------

Dim bExecutou As Boolean
sPosicaoErro = "ComplLanctonterno"

On Error GoTo Err_ComplementaLancamentoInterno

    
    Load LancamentoInterno
    With LancamentoInterno
        .SetParent Complementacao
        .SetPosition ((Screen.Width - .Width) / 2), 5300
        
        .Show vbModal, Me
    
        'Fecha Form
        Unload LancamentoInterno
        
        If .Alterou Then
            'Verificar, se houve inversão de imagem atualiza documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Lead1.Tag <> "0" Then
                bExecutou = G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , , Geral.Documento.Verso, Geral.Documento.Frente)
            Else
                bExecutou = True
            End If
            
            'Se Capa em duplicidade, Devolver docto com ocorrência após complementado
            If Geral.Capa.Duplicidade <> 0 Then
               If Not GeraOcorrenciaDocto Then bExecutou = False
            End If
            
            'Se Existe alteração de rotacionamento da imagem, salvar nova posiçao
            If bExecutou And iFlagRotacao <> 0 Then
                bExecutou = AlteraRotacao
            End If
            
            If Not bExecutou Then
                Beep
                MsgBox "Não foi possível complementar o documento!", vbCritical + vbOKOnly, App.Title
                iSituacao = 1: Exit Function
            End If
        Else
            iSituacao = 1: Exit Function
        End If
    End With

    '------------------------------------------
    '---    Atualiza Grid de Navegação      ---
    '------------------------------------------
    grdDocumentos.Col = iColTpDocto:    grdDocumentos.Text = etpdocLancamentoInterno
    If Geral.Capa.Duplicidade <> 0 Then
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = "D"
    Else
        grdDocumentos.Col = iColStatus:     grdDocumentos.Text = Geral.Documento.Status
    End If
    
    'Verifica se houve inversão de imagem
    If Lead1.Tag = "1" Then
        grdDocumentos.Col = iColFrente: grdDocumentos.Text = Geral.Documento.Verso
        grdDocumentos.Col = iColVerso:  grdDocumentos.Text = Geral.Documento.Frente
    End If
    
    'Finaliza complementação com sucesso
    iSituacao = 0: Exit Function

Err_ComplementaLancamentoInterno:
    
    'Finaliza complementação com Erro
    iSituacao = 2
    Unload LancamentoInterno
   
    Select Case TratamentoErro("1. Não foi possível complementar o documento.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select

End Function
Private Sub GeraOcorrenciaCapa()
    
On Error GoTo Err_GeraOcorrenciaCapa

    'Insere nova capa de envelope/malote tornando-o em complementação (Status = 2)
    With Modulo.qryAtualizaOcorrenciaCapa
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Capa.IdCapa       'Número do IdCapa
        .rdoParameters(3) = 999                     'Código de Ocorrência
        .Execute
        
        'Verifica se houve erro na atualização
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível realizar a ocorrência para esta capa.", vbCritical + vbOKOnly, App.Title
        End If
    End With
    
    Exit Sub
    
Err_GeraOcorrenciaCapa:
    
    Select Case TratamentoErro("Não foi possível realizar ocorrência para esta capa.", Err, rdoErrors, False)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Sub
Private Sub MotivoExclusao(ByVal iIdCapa As Integer, ByVal sMotivo As String)

On Error GoTo Err_MotivoExclusao

    'Gravar Motivo de Exclusão da Capa
    With Modulo.qryInsereMotivoExclusao
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = iIdCapa                     'IdCapa
        .rdoParameters(3) = sMotivo                     'MotivoExclusao
        .Execute
    End With

    If Modulo.qryInsereMotivoExclusao(0).Value = 1 Then
        MsgBox "Ocorreu um erro ao inserir motivo de exclusão.", vbInformation, App.Title
        Exit Sub
    End If

Exit Sub

Err_MotivoExclusao:
    Select Case TratamentoErro("Não foi possível inserir motivo de exclusão.", Err, rdoErrors, False)
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select

End Sub
