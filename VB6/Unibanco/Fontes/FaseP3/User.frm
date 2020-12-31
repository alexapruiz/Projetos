VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form Password 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2124
   ClientLeft      =   3348
   ClientTop       =   3252
   ClientWidth     =   4488
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "User.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2124
   ScaleWidth      =   4488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Backup"
      Height          =   372
      Index           =   1
      Left            =   1572
      TabIndex        =   4
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3024
      TabIndex        =   5
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Atual"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1332
   End
   Begin VB.PictureBox Panel3D1 
      AutoSize        =   -1  'True
      Height          =   1452
      Left            =   120
      ScaleHeight     =   1404
      ScaleWidth      =   4164
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4212
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   384
         Left            =   300
         Picture         =   "User.frx":030A
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   9
         Top             =   240
         Width           =   384
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         Left            =   2112
         MaxLength       =   10
         TabIndex        =   0
         Top             =   60
         Width           =   1584
      End
      Begin VB.TextBox TxtSenha 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         IMEMode         =   3  'DISABLE
         Left            =   2112
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   492
         Width           =   1584
      End
      Begin DATEEDITLib.DateEdit txtDataProcessamento 
         Height          =   384
         Left            =   2112
         TabIndex        =   2
         Top             =   900
         Width           =   1584
         _Version        =   65537
         _ExtentX        =   2794
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.Label Label3 
         Caption         =   "Data Movimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   84
         TabIndex        =   10
         Top             =   924
         Width           =   2028
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   900
         TabIndex        =   8
         Top             =   480
         Width           =   864
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   336
         Left            =   900
         TabIndex        =   7
         Top             =   120
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *************************************
' * Definição das Variáveis do Módulo *
' *************************************
Public Cancelou As Boolean          ' Indicação de Cancelamento
Public SenhaOk As Boolean           ' Indicação de Senha Digitada
Private qryUsuario  As rdoQuery     ' Chamada de Store Procedure para Leitura de Usuário
Private qrydataback As rdoQuery
Private tbUsuario As rdoResultset    ' Leitura do Usuário
    
Dim strServidor         As String
Dim strDataBaseAtual    As String
Dim strDataBaseBackup   As String
Dim strUsuario          As String
Dim strSenha            As String
Public bSenhaAlterada   As Boolean

' ********************************************
' * Efetua a verificação do Login do Usuário *
' ********************************************
Private Sub cmdConfirma_Click(Index As Integer)

Dim bBaseBackup     As Boolean
'    Dim Servidor        As String
'    Dim DataBaseAtual   As String
'    Dim DataBaseBackup  As String
'    Dim Usuario         As String
'    Dim Senha           As String
Dim Userlogin       As String
Dim USerSenha       As String
    
Dim eRetorno        As enumRetornoUsuario
    
    Cancelou = False
    SenhaOk = False

    ' ******************************************
    ' * Testa Digitação Obrigatória do Usuário *
    ' ******************************************
    If Trim(txtUsuario.Text) = "" Then
        Beep
        MsgBox "Digite o Usuário !", vbExclamation + vbOKOnly, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
    ' ****************************************
    ' * Testa Digitação Obrigatória da Senha *
    ' ****************************************
    If Trim(TxtSenha.Text) = "" Then
        Beep
        MsgBox "Digite a Senha !", vbExclamation + vbOKOnly, App.Title
        With TxtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
    ' ****************************************
    ' * Testa Digitação Obrigatória da Data  *
    ' ****************************************
    If Trim(txtDataProcessamento.Text) = "" Then
        Beep
        MsgBox "Digite a Data !", vbExclamation + vbOKOnly, App.Title
        With txtDataProcessamento
            .SetFocus
        End With
        Exit Sub
    End If
    
    If DataOk(Val(txtDataProcessamento.Text)) Then
        Geral.DataProcessamento = DataAAAAMMDD(Val(txtDataProcessamento.Text))
        'Geral.DiretorioImagens = PegarOpcaoINI("Diretorios", "Imagens", App.Path & "\IMAGENS") & "\" & DataAAAAMMDD(Val(txtDataProcessamento.Text)) & "\"
    Else
        MsgBox "A data informada não é válida!" & vbCr & "Obs.: O ano deve ser maior que 1997 e menor que 2051!", vbExclamation + vbOKOnly, App.Title
        txtDataProcessamento.SetFocus
        Exit Sub
    End If

    On Error GoTo ErroLogin

    Geral.Backup = CBool(Index = 1)
    Principal.mnuAltDataMovimento(0).Visible = Geral.Backup

    '''''''''''''''''''''''''''
    ' Rotina de inicialização '
    '''''''''''''''''''''''''''
'    Servidor = PegarOpcaoINI("Conexao", "Servidor", App.path & "\MDI_Ubb.ini")
'    DataBaseAtual = PegarOpcaoINI("Conexao", "DataBaseAtual", App.path & "\MDI_Ubb.ini")
'    DataBaseBackup = PegarOpcaoINI("Conexao", "DataBaseBackup", App.path & "\MDI_Ubb.ini")
'    Usuario = PegarOpcaoINI("Conexao", "Usuario", App.path & "\MDI_Ubb.ini")
'    Senha = PegarOpcaoINI("Conexao", "Senha", App.path & "\MDI_Ubb.ini")
    
'    Servidor = strServidor
'    DataBaseAtual = strDataBaseAtual
'    DataBaseBackup = strDataBaseBackup
'    Usuario = strUsuario
'    Senha = strSenha

    If Index = 0 Then
        Geral.StringConexao = "driver={SQL Server};Server=" & strServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & strDataBaseAtual & ";provider=sqloledb"
        bBaseBackup = False
    Else
        Geral.StringConexao = "driver={SQL Server};server=" & strServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & strDataBaseBackup & ";provider=sqloledb"
        bBaseBackup = True
    End If

    With Geral.Banco
        .Connect = Geral.StringConexao
        .CursorDriver = rdUseClientBatch
        .EstablishConnection rdDriverNoPrompt
    End With
    
    'Verificação de Data de Processamento qdo Backup
    If bBaseBackup = True Then
        Userlogin = Trim(txtUsuario.Text)
        USerSenha = TxtSenha.Text
        If VerificaDataBackup = False Then
            Unload Me
            txtUsuario.Text = Userlogin
            TxtSenha.Text = USerSenha
            Exit Sub
        End If
    End If
    
    Set qryUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")
    With qryUsuario
        .rdoParameters(0).Value = Trim(txtUsuario.Text)
        Set tbUsuario = .OpenResultset(rdConcurReadOnly)
    End With
    
     'Obtem o Nome do usuário
     If Not tbUsuario.EOF Then
            Geral.NomeUsuario = Trim(tbUsuario!Nome)
            'Verifica se existe última data de troca se senha por usuario
            If IsNull(tbUsuario!DataUltimaTrocaSenha) Then
                Geral.DataUltimaTrocaSenhaUsuario = 19900101   'Se data = null, força data antiga
            Else
                Geral.DataUltimaTrocaSenhaUsuario = Format(tbUsuario!DataUltimaTrocaSenha, "yyyymmdd")
            End If
          'Carrega todos grupos do usuário
          Geral.GrupoUsuario = ""
          While Not tbUsuario.EOF()
              Geral.GrupoUsuario = Geral.GrupoUsuario & UCase(Trim(tbUsuario!IdGrupo)) & "*"
              tbUsuario.MoveNext
          Wend
          tbUsuario.MoveFirst
     Else
          Geral.NomeUsuario = "Desenvolvimento"
          Geral.GrupoUsuario = "COO"
     End If
    
    ' ***********************************
    ' * Verificação do Login do Usuário *
    ' ***********************************
    eRetorno = VerificaUsuario(tbUsuario, txtUsuario.Text, TxtSenha.Text, bBaseBackup)
    
    If eRetorno = eSUPERVISOR Then
        Geral.idUsuario = 0
        SenhaOk = True
    ElseIf eRetorno = eNAO_EXISTENTE Then
        Beep
        MsgBox "Usuário não Cadastrado !", vbExclamation, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    ElseIf eRetorno = eSENHA_INCORRETA Then
        Beep
        MsgBox "Senha não Confere !", vbExclamation, App.Title
        With TxtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    Else
        SenhaOk = True
                    
        While Not tbUsuario.EOF
            '* Grava Informação de IdUsuario *'
            Geral.idUsuario = tbUsuario!idUsuario
            tbUsuario.MoveNext
        Wend
        
        Principal.MnuConRelValeTransporte(0).Enabled = bBaseBackup

        DoEvents
    End If
    
    If SenhaOk Then
        Geral.Usuario = Trim(txtUsuario)
        
        '*********************************************'
        '* Verifica se usuário esta ativo ou Inativo *'
        '* se usuário não for Desenv-Desenvolvimento *'
        '*********************************************'
        If UCase(Geral.Usuario) <> "DESENV" Then
            If Not DataUltimoLogon Then
                Unload Me
            Else
                'Verifica versao do sistema
                If Not VersaoCorreta Then
                    Unload Me
                End If
                'Carrega dias para forçar troca se senha
                If InStr(Geral.GrupoUsuario, "SPT") = 0 Then
                    If Not CarregaDiasTrocaSenha Then
                        Unload Me
                    Else
                        If SenhaUsuarioExpirada Then
                            bSenhaAlterada = False
                            AlteraSenha.Show vbModal, Me
                            If Not bSenhaAlterada Then
                                Unload Me
                            End If
                        End If
                    End If
                End If
                
            End If
        End If
        
        Me.Hide

    Else
        Geral.Banco.Close
    End If

    Exit Sub

ErroLogin:
    Select Case TratamentoErro("Erro na conexão com Banco de Dados.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
        .SetFocus
    End With

End Sub
' ******************************
' * Cancela o Login no Sistema *
' ******************************
Private Sub cmdSair_Click()
    Cancelou = True
    Me.Hide
End Sub
Private Sub Form_Activate()

    Cancelou = True

'    txtDataProcessamento.Text = Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000")
    'Obtem a ultima data de movimento contida na tabela parâmetro
    txtDataProcessamento.Text = UltimaDataMovimento

End Sub
' **************************************
' * Carrega Módulo de Login no Sistema *
' **************************************
Private Sub Form_Load()
    
    Cancelou = False
    SenhaOk = False
    Set Geral.Banco = New rdoConnection
End Sub

Private Sub txtDataProcessamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If DataOk(Val(txtDataProcessamento.Text)) Then
            cmdConfirma_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

' ************************************
' * Ajustando Seleção do Campo Senha *
' ************************************
Private Sub txtSenha_GotFocus()
    With TxtSenha
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

' **************************************
' * Ajustando Seleção do Campo Usuario *
' **************************************
Private Sub txtUsuario_GotFocus()
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Function VerificaDataBackup() As Boolean
'* Verifica se Data de Processamento existe no Backup*'

On Error GoTo ErroActive

    Dim qryCargaTabelas       As rdoQuery
    Dim qryUpdateCargaTabelas As rdoQuery
    Dim tb                   As rdoResultset
    
    VerificaDataBackup = True
    
    
    Set qryCargaTabelas = Geral.Banco.CreateQuery("", "{call GetCargaTabelas (?)}")
        qryCargaTabelas.rdoParameters(0) = Geral.DataProcessamento
    
        Set tb = qryCargaTabelas.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
        If Not tb.EOF And Not IsNull(tb!CargaTabelas) Then
        Else
            If InStr(1, Geral.StringConexao, "Backup", vbTextCompare) <> 0 Then
                MsgBox "Não existe esta Data de Movimento na base Backup.", _
                vbExclamation + vbOKOnly, App.Title
                VerificaDataBackup = False
            End If
        End If
        
Exit Function
ErroActive:
    Select Case TratamentoErro("Erro na verificação de Data de Movimento", Err, rdoErrors)
        Case vbCancel
            End
        Case vbRetry
            Resume
        End Select
        
End Function
Function DataUltimoLogon() As Boolean
'* Verifica se Usuário esta Ativo / Inativo e a qtde de dias que esta Inativo *'

On Error GoTo TrataErro

    Dim qryDataUltimoLogon  As rdoQuery     'Traz a Qtde de Dias que Usuário não se loga no Sistema
    Dim qryAtualizaData     As rdoQuery     'Atualiza com a data atual o último logon do usuário se ele estiver ativo
    Dim rsDataUltimoLogon   As rdoResultset 'Recordset
    Dim rsAtualizaData      As rdoResultset 'Recordset
    Dim ParamDiasInativo    As Long         'Qtde de Dias que usuário pode ficar inativo segundo a tabela de parâmetros
    Dim QtdesDiasInativo    As Long         'Qtde de Dias que usuário está Inativo

    Set qryDataUltimoLogon = Geral.Banco.CreateQuery("", "{Call GetDataUltimoLogon(?)}")
    Set qryAtualizaData = Geral.Banco.CreateQuery("", "{? =Call AtualizaDataUltimoLogon(?)}")

    With qryDataUltimoLogon
        .rdoParameters(0) = Geral.Usuario
        Set rsDataUltimoLogon = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not rsDataUltimoLogon.EOF Then
    
        '* Qtde de dias que usuário esta inativo
        QtdesDiasInativo = rsDataUltimoLogon!QtdeDiasInativo

        rsDataUltimoLogon.MoreResults

        '* Qtde de dias que usuário poderá ficar inativo
        ParamDiasInativo = rsDataUltimoLogon!Diasinativo

    End If

    '* Se a qtde de Dias Inativo for maior que a qtde máxima permitida de dias
    '* Inativo, usuário não poderá se logar no sistema,  sem  a autorização  do
    '* Supervisor ou do Suporte.
    If QtdesDiasInativo > ParamDiasInativo Then

        MsgBox "Usuário se encontra Inativo " & QtdesDiasInativo & " Movimentos.", vbExclamation + vbOKOnly, App.Title
        DataUltimoLogon = False

    Else
        '* Se Usuário estiver Ativo esta procedure irá atualizar sua última Data de Logon com getdate() *'
        With qryAtualizaData
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.Usuario
            Set rsAtualizaData = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            
            '* Se Parametro de Retorno for <> 0 ocorreu erro de atualização*'
            If .rdoParameters(0) <> 0 Then
                GoTo TrataErro
                DataUltimoLogon = False
            End If
        End With

        DataUltimoLogon = True

    End If
    
Exit Function

TrataErro:
    Select Case TratamentoErro("Não foi possível verificar se usuário está Ativo ou Inativo.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
End Select
End Function
Private Function UltimaDataMovimento() As String

Dim qryUltimaData   As New rdoQuery
Dim rsUltimaData    As rdoResultset
Dim strConexao      As String
Dim conBanco        As New rdoConnection

On Error GoTo Err_UltimaDataMovimento
    
    UltimaDataMovimento = ""
    
    '''''''''''''''''''''''''''
    ' Rotina de inicialização '
    '''''''''''''''''''''''''''
    strServidor = PegarOpcaoINI("Conexao", "Servidor", App.path & "\MDI_Ubb.ini")
    strDataBaseAtual = PegarOpcaoINI("Conexao", "DataBaseAtual", App.path & "\MDI_Ubb.ini")
    strDataBaseBackup = PegarOpcaoINI("Conexao", "DataBaseBackup", App.path & "\MDI_Ubb.ini")
    strUsuario = PegarOpcaoINI("Conexao", "Usuario", App.path & "\MDI_Ubb.ini")
    strSenha = PegarOpcaoINI("Conexao", "Senha", App.path & "\MDI_Ubb.ini")

    strConexao = "driver={SQL Server};Server=" & strServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & strDataBaseAtual & ";provider=sqloledb"

    With conBanco
        .Connect = strConexao
        .CursorDriver = rdUseClientBatch
        .EstablishConnection rdDriverNoPrompt
    End With
    
    
    Set qryUltimaData = conBanco.CreateQuery("", "{Call GetUltimaDataMovimento(?,?)}")

    qryUltimaData.rdoParameters(0).Direction = rdParamOutput    'Ultima data existente em parâmetro
    qryUltimaData.rdoParameters(1).Direction = rdParamOutput    'Data do servidor
    
    Set rsUltimaData = qryUltimaData.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If qryUltimaData.rdoParameters(0).Value <> 0 Then
        If qryUltimaData.rdoParameters("@DataMovimento").Value <= qryUltimaData.rdoParameters("@DataDoServidor").Value Then
            UltimaDataMovimento = Format(DataDDMMAAAA(qryUltimaData.rdoParameters("@DataMovimento").Value), "00/00/0000")
        End If
    End If
    

Err_UltimaDataMovimento:
    If Not (rsUltimaData Is Nothing) Then Set rsUltimaData = Nothing
    If Not (qryUltimaData Is Nothing) Then Set qryUltimaData = Nothing
    conBanco.Close
    
End Function
Private Function VersaoCorreta() As Boolean

Dim rsVersao    As rdoResultset

On Error GoTo Err_VersaoCorreta

    Set rsVersao = Geral.Banco.OpenResultset("Select * From MDI_Versao", rdOpenKeyset, rdConcurReadOnly)
    
    'Verifica versão do sistema
    If CStr(rsVersao!VersaoNumero) <> (App.Major & App.Minor & App.Revision) Then
        Beep
        MsgBox "Versão incorreta do sistema !" & vbCrLf & vbCrLf & "Favor entrar em contato com o depto. de suporte", vbCritical, App.Title
        Exit Function
    End If
    
    VersaoCorreta = True
    
    Exit Function
    
Err_VersaoCorreta:
    Beep
    MsgBox "Erro na verificação de Versão do sistema" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    
End Function
Private Function CarregaDiasTrocaSenha() As Boolean

Dim qryParametro    As New rdoQuery
Dim rsParametro     As rdoResultset

On Error GoTo Err_CarregaDiasTrocaSenha

    CarregaDiasTrocaSenha = False

    Set qryParametro = Geral.Banco.CreateQuery("", "{call LerParametroTrocaSenha(?)}")

    qryParametro.rdoParameters(0) = Geral.DataProcessamento
    Set rsParametro = qryParametro.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If rsParametro.EOF Then
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical, App.Title
        GoTo Exit_CarregaDiasTrocaSenha
    End If
    
    If rsParametro!QtdeDiasTrocaSenha = 0 Then
        MsgBox "Problema de informações na tabela Parâmetro, favor contatar o suporte." & vbCrLf & vbCrLf & _
                "Ocorrência: Dias para troca da senha não informada.", _
                vbCritical, App.Title
        GoTo Exit_CarregaDiasTrocaSenha
    End If
    
    Geral.QtdeDiasTrocaSenha = rsParametro!QtdeDiasTrocaSenha

    CarregaDiasTrocaSenha = True

Exit_CarregaDiasTrocaSenha:
    If Not (rsParametro Is Nothing) Then Set rsParametro = Nothing
    If Not (qryParametro Is Nothing) Then Set qryParametro = Nothing
    Exit Function
    
Err_CarregaDiasTrocaSenha:
    Beep
    MsgBox "Erro na verificação de Parâmetros do sistema" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    GoTo Exit_CarregaDiasTrocaSenha

End Function

Private Function SenhaUsuarioExpirada() As Boolean

Dim sysTime As SYSTEMTIME
Dim lDataSistema As Long
Dim dtDataDe, dtDataAte As Date
Dim iSomaData As Integer

SenhaUsuarioExpirada = True
    
    If Geral.DataUltimaTrocaSenhaUsuario = 19900101 Then Exit Function
    
    GetLocalTime sysTime

    lDataSistema = Val(Format(sysTime.wYear, "0000") & Format(sysTime.wMonth, "00") & Format(sysTime.wDay, "00"))

    If Geral.DataUltimaTrocaSenhaUsuario > lDataSistema Then
        Exit Function
    End If
    
    dtDataDe = CVDate(DataDD_MM_AAAA(Geral.DataUltimaTrocaSenhaUsuario))
    dtDataAte = CVDate(DataDD_MM_AAAA(lDataSistema))
    
    iSomaData = 0
    While dtDataDe < dtDataAte
        If Weekday(dtDataDe) >= 2 And Weekday(dtDataDe) <= 6 Then iSomaData = iSomaData + 1
        dtDataDe = dtDataDe + 1
    Wend
    
    If iSomaData <= Geral.QtdeDiasTrocaSenha Then
        SenhaUsuarioExpirada = False
    End If
    
End Function
