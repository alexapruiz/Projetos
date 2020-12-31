VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControleGeracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Geração"
   ClientHeight    =   3576
   ClientLeft      =   36
   ClientTop       =   288
   ClientWidth     =   4272
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3576
   ScaleWidth      =   4272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Identificação do lacre"
      Height          =   732
      Left            =   96
      TabIndex        =   13
      Top             =   0
      Width           =   4092
      Begin VB.TextBox txtNumLacre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   348
         Left            =   2592
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Lacre"
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
         Left            =   288
         TabIndex        =   14
         Top             =   288
         Width           =   1464
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   96
      TabIndex        =   11
      Top             =   2688
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   396
      Left            =   2205
      TabIndex        =   10
      Top             =   3072
      Width           =   972
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "&Gerar"
      Height          =   396
      Left            =   1095
      TabIndex        =   9
      Top             =   3072
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   1884
      Left            =   96
      TabIndex        =   5
      Top             =   720
      Width           =   4092
      Begin VB.TextBox txtRemessa 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   348
         Left            =   2592
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1212
      End
      Begin VB.TextBox txtDocumentos 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   348
         Left            =   2592
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1392
         Width           =   1212
      End
      Begin VB.TextBox txtCapas 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   348
         Left            =   2592
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1008
         Width           =   1212
      End
      Begin VB.TextBox txtLotes 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   348
         Left            =   2592
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   624
         Width           =   1212
      End
      Begin VB.Label lblRemessa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remessa"
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
         Left            =   264
         TabIndex        =   12
         Top             =   288
         Width           =   816
      End
      Begin VB.Label lblDoctos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos"
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
         Left            =   264
         TabIndex        =   8
         Top             =   1440
         Width           =   1092
      End
      Begin VB.Label lblCapas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capas"
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
         Left            =   264
         TabIndex        =   7
         Top             =   1056
         Width           =   552
      End
      Begin VB.Label lblLotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes"
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
         Left            =   264
         TabIndex        =   6
         Top             =   672
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmControleGeracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qryVerificaLotePendente     As rdoQuery
Private qryObtemLotesExportacao     As rdoQuery
Private qryObtemCapasLote           As rdoQuery
Private qryObtemLog                 As rdoQuery
Private qryObtemDocumentosCapa      As rdoQuery
Private qryObtemAgencias            As rdoQuery
Private qryTotalizaAgencia          As rdoQuery
Private qryEncerraMovimento         As rdoQuery

Private m_bRegerar                  As Boolean
'
'Rotina que soma na variavel origem.
'
'se enviado -1 zera o contador
'
Private Sub Sequencia(ByRef pVar As String)

    Static lVar        As Long
    
    If Val(pVar) = -1 Then lVar = 0: Exit Sub
    
    lVar = lVar + 1
    
    pVar = FormataString(lVar, "0", Len(pVar), True)

End Sub

'
'Funcao FormataString
'
'
Private Function FormataString( _
    ByVal pOque As Variant, _
    ByVal pCompletarCom As Variant, _
    ByVal pFieldLen As Integer, _
    ByVal pAEsquerda As Boolean) As Variant
    
    If pFieldLen <= 0 Then FormataString = pOque: Exit Function
    If pCompletarCom = "" Then FormataString = pOque: Exit Function
    If pFieldLen < Len(pOque) Then FormataString = pOque: Exit Function

    If pAEsquerda Then
        FormataString = Right(String(pFieldLen - Len(pOque), pCompletarCom) & pOque, pFieldLen)
    Else
        FormataString = Left(pOque & String(pFieldLen - Len(pOque), pCompletarCom), pFieldLen)
    End If
    
End Function

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdGerar_Click()

    Dim rsAgencia       As rdo.rdoResultset
    Dim rsLote          As rdo.rdoResultset
    Dim rsCapa          As rdo.rdoResultset
    Dim rsCapaLog       As rdo.rdoResultset
    Dim rsDocto         As rdo.rdoResultset
    Dim rsDoctoLog      As rdo.rdoResultset
    
    Dim Dados           As cg_DADOS
    Dim agencia         As cg_AGENCIA
    Dim CD              As cg_CD
    Dim iFile           As Integer
    Dim iP              As Integer
    Dim lC              As Long
    Dim Origem          As String
    Dim Destino         As String
    
    On Error GoTo Erro_Geracao:
    
    ProgressBar1.Value = 0
    
    txtRemessa.Text = ""
    txtLotes.Text = ""
    txtCapas.Text = ""
    txtDocumentos.Text = ""
    
    Me.Refresh
    
    If (Trim(txtNumLacre.Text) = "") Or (Not IsNumeric(txtNumLacre.Text)) Then
        MsgBox "Número do lacre inválido.", vbExclamation
        txtNumLacre.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Após a geração dos dados não será permitida a captura de novos documentos nesta data." & Chr(10) & "Confirma a geração dos dados ?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    If m_bRegerar Then
        If Not ShellDelete(Geral.CDR.Drive & "*.*") Then
            MsgBox "Não foi possível limpar o CD.", vbCritical
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    'Cria diretorio do cd caso não exista'
    ''''''''''''''''''''''''''''''''''''''
    If Not DirExists(Geral.DiretorioDados & DIR_CD) Then
        '''''''''''''''''''''
        'Cria sub-diretórios'
        '''''''''''''''''''''
        If Not CriaDir(Geral.DiretorioDados & DIR_CD) Then
            MsgBox "Erro ao tentar criar o diretório '" & Geral.DiretorioDados & DIR_CD & "'", vbCritical
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''
    'Verificacao de lote pendente'
    ''''''''''''''''''''''''''''''
    With qryVerificaLotePendente
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .Execute
        
        If .rdoParameters(0) = 1 Then
            MsgBox "Não foi possível concluir a geração de arquivos." & Chr(10) & "Existe lote pendente de liberação no Controle de Qualidade.", vbExclamation
            Exit Sub
        End If
    End With
    
    '''''''''''''''''''''''''''''''''''
    'Obtem os lotes a serem exportados'
    '''''''''''''''''''''''''''''''''''
    With qryObtemLotesExportacao
        .rdoParameters(0) = Geral.DataProcessamento
        Set rsLote = qryObtemLotesExportacao.OpenResultset(rdOpenStatic, rdConcurReadOnly)
        
        If rsLote.EOF Then
            MsgBox "Não existem lotes a serem exportados.", vbExclamation
            rsLote.Close
            '''''''''''''''''''''''''''
            'Encerramento do Movimento'
            '''''''''''''''''''''''''''
            With qryEncerraMovimento
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento
                .Execute
                If .rdoParameters(0) <> 0 Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Erro no encerramento do movimento.", vbCritical
                End If
            End With
            Exit Sub
        End If
    End With
    
    '''''''''''''''''''''''''''''''''''''''''
    'Exclui os arquivos que ja foram gerados'
    '''''''''''''''''''''''''''''''''''''''''
    If FileExist(Geral.DiretorioDados & DIR_CD & ARQ_DADOS) Then ShellDelete Geral.DiretorioDados & DIR_CD & ARQ_DADOS
    If FileExist(Geral.DiretorioDados & DIR_CD & ARQ_CD) Then ShellDelete Geral.DiretorioDados & DIR_CD & ARQ_CD
    If FileExist(Geral.DiretorioDados & DIR_CD & ARQ_AGENCIA) Then ShellDelete Geral.DiretorioDados & DIR_CD & ARQ_AGENCIA
    
    
    txtRemessa.Text = "00001"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                              INICIALIZACAO DO ARQUIVO DADOS.DAT                                    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Screen.MousePointer = vbHourglass
    
    
    
    iFile = FreeFile
    Open Geral.DiretorioDados & DIR_CD & ARQ_DADOS For Binary As iFile
    
    ''''''''''''''''''''
    'Reseta a sequencia'
    ''''''''''''''''''''
    Sequencia -1
    With Dados.Header
        Sequencia .Sequencial
        .TipoRegistro = "H"
        .DataProcessamento = CStr(Geral.DataProcessamento)
        .AgOrig = Geral.AgenciaApresentante
        .Remessa = "00001" 'Fixo 1, enquanto existir uma remessa por dia
        .CrLf = vbCrLf
    End With
    
    '''''''''''''''''''''''''''
    'Grava o header no arquivo'
    '''''''''''''''''''''''''''
    Put #iFile, , Dados.Header
    
    
    ''''''''''''''''
    'Loop dos lotes'
    ''''''''''''''''
    lC = rsLote.RowCount
    txtLotes.Text = lC
    
    
    Do While Not rsLote.EOF()
    
        iP = Abs(rsLote.AbsolutePosition / lC) * 100
        ProgressBar1.Value = iP
        
        ''''''''''''''''''''''''''''''
        'Gravação de lote
        ''''''''''''''''''''''''''''''
        With Dados.Lote
            .TipoRegistro = "L"
            Sequencia .Sequencial
            .IdLote = FormataString(rsLote!IdLote, "0", Len(.IdLote), True)
            .Status = rsLote!Status
            .Prioridade = CStr(rsLote!Prioridade)
            .CrLf = vbCrLf
        End With
        '''''''''''''''''''''''''
        'Grava o lote no arquivo'
        '''''''''''''''''''''''''
        Put #iFile, , Dados.Lote
    
        With qryObtemCapasLote
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = rsLote!IdLote
            Set rsCapa = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
        End With
        
        txtCapas.Text = Val(txtCapas.Text) + rsCapa.RowCount
        
        ''''''''''''''''
        'Loop das capas'
        ''''''''''''''''
        Do While Not rsCapa.EOF()
            ''''''''''''''''''''''''''''''''''''
            'Gravar registro da capa no arquivo'
            ''''''''''''''''''''''''''''''''''''
            With Dados.Capa
                .TipoRegistro = "C"
                Sequencia .Sequencial
                .IdLote = FormataString(rsCapa!IdLote, "0", Len(.IdLote), True)
                .IdEnv_Mal = rsCapa!IdEnv_Mal
                .Capa = FormataString(rsCapa!Capa, "0", Len(.Capa), True)
                .Num_Malote = FormataString(rsCapa!Num_Malote, "0", Len(.Num_Malote), True)
                .AgOrig = Geral.AgenciaApresentante
                .Status = rsCapa!Status
                .Ocorrencia = FormataString(rsCapa!Ocorrencia, "0", Len(.Ocorrencia), True)
                .Duplicidade = rsCapa!Duplicidade
                .CrLf = vbCrLf
            End With
            Put #iFile, , Dados.Capa
        
        
            '''''''''''''''''''''
            'Obtem log das capas'
            '''''''''''''''''''''
            With qryObtemLog
                .rdoParameters(0) = Geral.DataProcessamento
                .rdoParameters(1) = rsCapa!IdCapa
                .rdoParameters(2) = 0
                Set rsCapaLog = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
            End With
            '''''''''''''''''''''''
            'Loop do log das capas'
            '''''''''''''''''''''''
            Do While Not rsCapaLog.EOF
                With Dados.Log
                    .TipoRegistro = "G"
                    Sequencia .Sequencial
                    .Data = FormataString(Format(rsCapaLog!Data, "yyyymmdd hh:mm:ss"), " ", Len(.Data), False)
                    .Login = FormataString(rsCapaLog!Login, " ", Len(.Login), False)
                    .Acao = FormataString(rsCapaLog!Acao, "0", Len(.Acao), True)
                    .CrLf = vbCrLf
                End With
                Put #iFile, , Dados.Log
                
                rsCapaLog.MoveNext
            Loop
            ''''''''''''''''''''''''''
            'Obtem documentos da capa'
            ''''''''''''''''''''''''''
            With qryObtemDocumentosCapa
                .rdoParameters(0) = Geral.DataProcessamento
                .rdoParameters(1) = rsCapa!IdCapa
                Set rsDocto = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
            End With
            
            txtDocumentos.Text = Val(txtDocumentos.Text) + rsDocto.RowCount
            
            Do While Not rsDocto.EOF()
                ''''''''''''''''''''''''''''''''''
                'Gravação do documento no arquivo'
                ''''''''''''''''''''''''''''''''''
                With Dados.Docto
                    .TipoRegistro = "D"
                    Sequencia .Sequencial
                    .OrdemCaptura = FormataString(rsDocto!OrdemCaptura, "0", Len(.OrdemCaptura), True)
                    .TipoDocto = FormataString(rsDocto!TipoDocto, "0", Len(.TipoDocto), True)
                    .Leitura = FormataString(rsDocto!Leitura, " ", Len(.Leitura), False)
                    .Frente = FormataString(rsDocto!Frente, " ", Len(.Frente), False)
                    .Verso = FormataString(rsDocto!Verso, " ", Len(.Verso), False)
                    .Status = rsDocto!Status
                    .Ordem = rsDocto!Ordem
                    .CrLf = vbCrLf
                End With
                Put #iFile, , Dados.Docto
                ''''''''''''''''''''''''
                'Obtem log do documento'
                ''''''''''''''''''''''''
                With qryObtemLog
                    .rdoParameters(0) = Geral.DataProcessamento
                    .rdoParameters(1) = rsCapa!IdCapa
                    .rdoParameters(2) = rsDocto!IdDocto
                    Set rsDoctoLog = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
                End With
                
                Do While Not rsDoctoLog.EOF()
                    ''''''''''''''''''''''''''''''
                    'Gravação do log do documento'
                    ''''''''''''''''''''''''''''''
                    With Dados.Log
                        .TipoRegistro = "G"
                        Sequencia .Sequencial
                        .Data = Format(rsDoctoLog!Data, "yyyymmdd hh:mm:ss")
                        .Login = FormataString(rsDoctoLog!Login, " ", Len(.Login), False)
                        .Acao = FormataString(rsDoctoLog!Acao, "0", Len(.Acao), True)
                        .CrLf = vbCrLf
                    End With
                    Put #iFile, , Dados.Log
                    rsDoctoLog.MoveNext
                Loop
                rsDocto.MoveNext
            Loop
        
            rsCapa.MoveNext
        Loop
    
        rsLote.MoveNext
    Loop
    
    ProgressBar1.Refresh
    
    With Dados.Trailler
        .TipoRegistro = "T"
        Sequencia .Sequencial
        .DataProcessamento = Geral.DataProcessamento
        .AgOrig = FormataString(Geral.AgenciaApresentante, "0", Len(.AgOrig), True)
        .Remessa = "00001"
    End With
    Put #iFile, , Dados.Trailler
    
    rsLote.Close
    rsCapa.Close
    rsCapaLog.Close
    rsDocto.Close
    rsDoctoLog.Close
    Close #iFile
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                            INICIALIZACAO DO ARQUIVO CD.ID                                          '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    iFile = FreeFile
    Open Geral.DiretorioDados & DIR_CD & ARQ_CD For Binary As #iFile
    
    With CD
        .agencia = FormataString(Geral.AgenciaApresentante, "0", Len(.agencia), True)
        .Data = FormataString(Geral.DataProcessamento, "0", Len(.Data), True)
        .Hora = Format(Now, "HH:MM:SS")
        .Remessa = "00001"
        .Numero_CD = "01"
    End With
    Put #iFile, , CD
    
    Close #iFile
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                            INICIALIZACAO DO ARQUIVO AGENCIA.DAT                                    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Iniciar transacao somente no processo de gravacao da agencia'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Geral.Banco.BeginTrans
    
    ''''''''''''''''''
    'Totaliza Agencia'
    ''''''''''''''''''
    With qryTotalizaAgencia
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.AgenciaApresentante
        .rdoParameters(2) = txtNumLacre
        .Execute
        If .rdoParameters(0) <> 0 Then
            Geral.Banco.RollbackTrans
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível concluir a geração." & Chr(10) & "Erro na totalização da Agência.", vbCritical
            ProgressBar1.Value = 0
            Exit Sub
        End If
    End With
    
    '''''''''''''''''''
    'Obtem as agencias'
    '''''''''''''''''''
    With qryObtemAgencias
        Set rsAgencia = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
    End With
    
    iFile = FreeFile
    Open Geral.DiretorioDados & DIR_CD & ARQ_AGENCIA For Binary As iFile
    
    '''''''''''''''''''''
    'Reseta o sequencial'
    '''''''''''''''''''''
    Sequencia -1
    With agencia.Header
        .TipoRegistro = "H"
        Sequencia .Sequencial
        .DataProcessamento = CStr(Geral.DataProcessamento)
        .AgOrig = Geral.AgenciaApresentante
        .Remessa = "00001"
        .CrLf = vbCrLf
    End With
    Put #iFile, , agencia.Header

    '''''''''''''''''''
    'Loop das agencias'
    '''''''''''''''''''
    Do While Not rsAgencia.EOF()
    
        With agencia.Registro
            .TipoRegistro = "R"
            Sequencia .Sequencial
            .agencia = FormataString(rsAgencia!agencia, "0", Len(.agencia), True)
            '''''''''''''''''''''''''''''''''''''''''''
            'Será necessário update do lacre na base ?'
            '''''''''''''''''''''''''''''''''''''''''''
            .Lacre = FormataString(rsAgencia!Lacre, "0", Len(.Lacre), True)
            .QtdInformada = FormataString(rsAgencia!QtdInformada, "0", Len(.QtdInformada), True)
            .HoraCadastrada = FormataString(rsAgencia!HoraCadastrada, " ", Len(.HoraCadastrada), False)
            .IdEnv_Mal = rsAgencia!IdEnv_Mal
            .CrLf = vbCrLf
        End With
        Put iFile, , agencia.Registro
        
        rsAgencia.MoveNext
    Loop
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Gravacao do Trailler do arquivo da agencia.dat'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    With agencia.Trailler
        .TipoRegistro = "T"
        Sequencia .Sequencial
        .DataProcessamento = CStr(Geral.DataProcessamento)
        .AgOrig = Geral.AgenciaApresentante
        .Remessa = "00001"
    End With
    Put #iFile, , agencia.Trailler
    '''''''''''''''''''''''''
    'Finalizacao da gravacao'
    '''''''''''''''''''''''''
    Close iFile
    
    
    '''''''''''''''''''''''''''
    'Encerramento do Movimento'
    '''''''''''''''''''''''''''
    With qryEncerraMovimento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .Execute
        If .rdoParameters(0) <> 0 Then
            Geral.Banco.RollbackTrans
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível concluir a geração." & Chr(10) & "Erro no encerramento do movimento.", vbCritical
            ProgressBar1.Value = 0
            Exit Sub
        End If
    End With
    
    Geral.Banco.CommitTrans
    
    ''''''''''''''''''''''''''''''''''''''
    'Verifica se o diretório no CD existe'
    ''''''''''''''''''''''''''''''''''''''
    Do While (DirExists(Geral.CDR.Drive & _
                     Geral.CDR.DiretorioDados & _
                     Geral.DataProcessamento & "\") = 0)
        '''''''''''''''''''''
        'Se não existe. CRIA'
        '''''''''''''''''''''
        If Not CriaDir(Geral.CDR.Drive & _
                       Geral.CDR.DiretorioDados & _
                       Geral.DataProcessamento & "\") Then
            '''''''''''''''''''''''''''''''''''''''''''
            'Se não conseguiu criar, não faz mais nada'
            '''''''''''''''''''''''''''''''''''''''''''
            If MsgBox("Não foi possível criar o diretório de dados no CD.", vbExclamation + vbRetryCancel) = vbCancel Then
                Screen.MousePointer = vbDefault
                ProgressBar1.Value = 0
                MsgBox "Não foi possível concluir a geração dos arquivos.", vbExclamation
                Exit Sub
            End If
        End If
    Loop
    
    ''''''''''''''''''''''''''''''''''
    'Tenta fazer a cópia dos arquivos'
    ''''''''''''''''''''''''''''''''''
    If Not ShellCopy(Geral.DiretorioDados & DIR_CD & "*.*", _
                     Geral.CDR.Drive & Geral.CDR.DiretorioDados & Geral.DataProcessamento & "\") Then
        Screen.MousePointer = vbDefault
        ProgressBar1.Value = 0
        Exit Sub
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''
    'Se é para copiar novamente as imagens'
    '''''''''''''''''''''''''''''''''''''''
    If m_bRegerar Then
    
        Do While (DirExists(Geral.CDR.Drive & _
                         Geral.CDR.DiretorioImagens & _
                         Geral.DataProcessamento & "\") = 0)
            '''''''''''''''''''''
            'Se não existe. CRIA'
            '''''''''''''''''''''
            If Not CriaDir(Geral.CDR.Drive & _
                           Geral.CDR.DiretorioImagens & _
                           Geral.DataProcessamento & "\") Then
                '''''''''''''''''''''''''''''''''''''''''''
                'Se não conseguiu criar, não faz mais nada'
                '''''''''''''''''''''''''''''''''''''''''''
                If MsgBox("Não foi possível criar o diretório de imagens no CD.", vbExclamation + vbRetryCancel) = vbCancel Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Não foi possível concluir a geração dos arquivos.", vbExclamation
                    ProgressBar1.Value = 0
                    Exit Sub
                End If
            End If
        Loop
        
        
        
        
        Origem = Geral.DiretorioImagens & "*.*"
        Destino = Geral.CDR.Drive & Geral.CDR.DiretorioImagens & Geral.DataProcessamento
    
        If Not ShellCopy(Origem, Destino) Then
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível copiar as imagens para o CD.", vbCritical
            ProgressBar1.Value = 0
            Exit Sub
        End If
    End If
    
    
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Geração completada com sucesso.", vbInformation
    ProgressBar1.Value = 0
    cmdGerar.Enabled = False
    
    Exit Sub
    
Erro_Geracao:
    Screen.MousePointer = vbDefault
    
    Geral.Banco.RollbackTrans
    ProgressBar1.Value = 0
    
    Close iFile
    
    Call TratamentoErro(Geral.Banco, "Erro na exportação.", Err, rdoErrors)

End Sub

Public Function getPerformance(mStart As Boolean) As String
    Static start    As Double
    Dim finish      As Double
    
    If mStart Then
        start# = Timer
        getPerformance = ""
    Else
        finish# = Timer
        getPerformance = Format$(finish# - start#, "##.######") & " secs."
    End If

End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdGerar_Click
    End If
End Sub

Private Sub Form_Load()

    Dim rsAgencias      As rdo.rdoResultset
    
    
    On Error GoTo Erro_Load

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'cria query de verificacao da existe de lote pendente'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set qryVerificaLotePendente = Geral.Banco.CreateQuery("", "{? = Call MDIAG_VerificaLotePendente(?)}")
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'cria query de seleção de lotes à serem exportados'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Set qryObtemLotesExportacao = Geral.Banco.CreateQuery("", "{Call MDIAG_ObtemLotesExportacao(?)}")
    '''''''''''''''''''''''''''''''''''''''''
    'cria query de seleção de capas por lote'
    '''''''''''''''''''''''''''''''''''''''''
    Set qryObtemCapasLote = Geral.Banco.CreateQuery("", "{Call MDIAG_ObtemCapasLote(?,?)}")
    ''''''''''''''''''''''''''''''
    'cria query de seleção de Log'
    ''''''''''''''''''''''''''''''
    Set qryObtemLog = Geral.Banco.CreateQuery("", "{Call MDIAG_ObtemLog(?,?,?)}")
    ''''''''''''''''''''''''''''''''''''
    'cria query de seleção de documento'
    ''''''''''''''''''''''''''''''''''''
    Set qryObtemDocumentosCapa = Geral.Banco.CreateQuery("", "{Call MDIAG_ObtemDocumentosCapa(?,?)}")
    ''''''''''''''''''''''''''''''''''''
    'cria query de seleção das agencias'
    ''''''''''''''''''''''''''''''''''''
    Set qryObtemAgencias = Geral.Banco.CreateQuery("", "{Call MDIAG_ObtemAgencias}")
    ''''''''''''''''''''''''''''''''''''''
    'cria query de totalizacao da agencia'
    ''''''''''''''''''''''''''''''''''''''
    Set qryTotalizaAgencia = Geral.Banco.CreateQuery("", "{? = Call MDIAG_TotalizaAgencia(?,?)}")
    ''''''''''''''''''''''''''''''''''''
    'cria query que encerra o movimento'
    ''''''''''''''''''''''''''''''''''''
    Set qryEncerraMovimento = Geral.Banco.CreateQuery("", "{? = Call MDIAG_EncerraMovimento(?)}")
    
    
    Set rsAgencias = qryObtemAgencias.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    m_bRegerar = False
    
    If Not rsAgencias.EOF() Then
        txtNumLacre.Text = rsAgencias!Lacre
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se entrar neste if é porque entrou para fazer uma nova '
        'geração dos dados portanto precisa fazer uma nova copia'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_bRegerar = True
        cmdGerar.Caption = "Regerar"
    End If
    
    rsAgencias.Close
    
    Exit Sub
    
Erro_Load:

    Call TratamentoErro(Geral.Banco, "Erro na abertura do " & Me.Caption & ".", Err, rdoErrors)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    qryVerificaLotePendente.Close
    qryObtemLotesExportacao.Close
    qryObtemCapasLote.Close
    qryObtemLog.Close
    qryObtemDocumentosCapa.Close
    qryObtemAgencias.Close
    qryTotalizaAgencia.Close
    qryEncerraMovimento.Close
End Sub


Private Sub txtCapas_GotFocus()
    SelecionarTexto txtCapas
End Sub


Private Sub txtDocumentos_GotFocus()
    SelecionarTexto txtDocumentos
End Sub


Private Sub txtLotes_GotFocus()
    SelecionarTexto txtLotes
End Sub


Private Sub txtNumLacre_GotFocus()
    SelecionarTexto txtNumLacre
End Sub

Private Sub txtNumLacre_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub

Private Sub txtRemessa_GotFocus()
    SelecionarTexto txtRemessa
End Sub


