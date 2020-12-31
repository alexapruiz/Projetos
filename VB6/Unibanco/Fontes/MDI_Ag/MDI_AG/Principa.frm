VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8544
   ClientLeft      =   48
   ClientTop       =   636
   ClientWidth     =   12228
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8544
   ScaleWidth      =   12228
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar BarMain 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   8208
      Width           =   12228
      _ExtentX        =   21569
      _ExtentY        =   593
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17622
            MinWidth        =   17622
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "21/12/00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "11:18"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuRecepcao 
      Caption         =   "&Recepção"
      Begin VB.Menu mnuRecRecepcao 
         Caption         =   "&Recepção..."
      End
      Begin VB.Menu mnuRecRegistroOcorrencia 
         Caption         =   "Registro de &Ocorrência..."
      End
   End
   Begin VB.Menu mnuCaptura 
      Caption         =   "&Captura"
      Begin VB.Menu mnuCapCaptura 
         Caption         =   "&Captura..."
      End
      Begin VB.Menu mnuCapCtrQualidade 
         Caption         =   "Controle de &Qualidade..."
      End
   End
   Begin VB.Menu mnuEstatistica 
      Caption         =   "&Estatística..."
   End
   Begin VB.Menu mnuControleGeracao 
      Caption         =   "Controle de &Geração..."
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Paint()
    Dim i As Long
    Dim iTam As Long
    
    Dim y As Single
    
    iTam = ScaleHeight / 256
    
    y = 0
    
    For i = 0 To 255
        Line (0, y)-(ScaleWidth, y + iTam), RGB(i, i, i), BF
        y = y + iTam
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iRet As Long
    On Error Resume Next
    If Geral.Scanner = escnVIPS Then
        ObScanner.Done
        Set ObScanner = Nothing
    End If

    Geral.Banco.Close
End Sub
Private Sub mnuCapCaptura_Click()
    Dim Captura As New Captura
    
    Screen.MousePointer = vbHourglass
    
    If Not Captura.SetConnection(Geral.Banco) Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
    Else
        If Not Captura.SetScanner(ObScanner) Then
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível estabelecer comunicação com o scanner.", vbExclamation
        Else
            Captura.SetAgApresentante Geral.AgenciaApresentante
            Captura.SetDataProcessamento Geral.DataProcessamento
            Captura.SetDirDados Geral.DiretorioDados
            Captura.SetDirImagens Geral.DiretorioImagens
            Captura.SetDirTrabalho Geral.DiretorioTrabalho
            Captura.SetTipoScanner Geral.Scanner
            Captura.SetUsuario Geral.Usuario.Login
            Screen.MousePointer = vbDefault
            Captura.ShowModal
        End If
    End If
End Sub

Private Sub mnuCapCtrQualidade_Click()

    Dim CtrlQualidade       As New ControleQualidade

    Screen.MousePointer = vbHourglass

    If (CtrlQualidade.SetConnection(Geral.Banco) And _
        CtrlQualidade.SetDataProcessamento(Geral.DataProcessamento) And _
        CtrlQualidade.SetUsuario(Geral.Usuario.Login) And _
        CtrlQualidade.SetIntervalo(Geral.Intervalo) And _
        CtrlQualidade.SetAtualizacao(Geral.Atualizacao) And _
        CtrlQualidade.SetDirImagens(Geral.DiretorioImagens) And _
        CtrlQualidade.SetDriveCDR(Geral.CDR.Drive) And _
        CtrlQualidade.SetAgenciaApresentante(Geral.AgenciaApresentante) And _
        CtrlQualidade.SetDirImagensCDR(Geral.CDR.DiretorioImagens)) Then

        Screen.MousePointer = vbDefault
        CtrlQualidade.ShowModal

    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
    End If

    Set CtrlQualidade = Nothing

End Sub

Private Sub mnuControleGeracao_Click()

    Dim ControleGeracao     As New ControleGeracao
    Dim rsParametro         As rdo.rdoResultset
    
    Screen.MousePointer = vbHourglass
    
    If (ControleGeracao.SetConnection(Geral.Banco) And _
        ControleGeracao.SetDataProcessamento(Geral.DataProcessamento) And _
        ControleGeracao.SetDiretorioDados(Geral.DiretorioDados) And _
        ControleGeracao.SetAgenciaApresentante(Geral.AgenciaApresentante) And _
        ControleGeracao.SetDirImagensCDR(Geral.CDR.DiretorioImagens) And _
        ControleGeracao.SetDriveCDR(Geral.CDR.Drive) And _
        ControleGeracao.SetDiretorioImagens(Geral.DiretorioImagens) And _
        ControleGeracao.SetDirDadosCDR(Geral.CDR.DiretorioDados)) Then

        Screen.MousePointer = vbDefault
        ControleGeracao.ShowModal
        Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call MDIAG_LerParametro}")

        With Geral.qryLeituraParametro
            Set rsParametro = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        If Not rsParametro.EOF() Then
            If Not IsNull(rsParametro!Hm_Fechamento) Then
                ''''''''''''''''''''''
                'Desebilitar os menus'
                ''''''''''''''''''''''
                Principal.mnuRecepcao.Enabled = False
                Principal.mnuCaptura.Enabled = False
                'Principal.mnuEstatistica.Enabled = False
                Principal.mnuControleGeracao.Enabled = True
            End If
        End If
        
        rsParametro.Close
        Geral.qryLeituraParametro.Close
        
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
    End If

End Sub

Private Sub mnuEstatistica_Click()
    Dim Estatistica As New Estatistica
    
    Screen.MousePointer = vbHourglass
    
    If Not Estatistica.SetConnection(Geral.Banco) Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
    Else
        Estatistica.SetAtualizacao Geral.Atualizacao
        Estatistica.SetDataProcessamento Geral.DataProcessamento
        Screen.MousePointer = vbDefault
        Estatistica.ShowModal
    End If
End Sub

Private Sub mnuRecRecepcao_Click()
    Dim Recepcao    As New MDI_Recepcao.Recepcao
    
    Screen.MousePointer = vbHourglass
    
    If Not Recepcao.SetConnection(Geral.Banco) Then
        MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
    Else
        If (Recepcao.SetAgenciaApresentante(Geral.AgenciaApresentante) And _
            Recepcao.SetDataProcessamento(Geral.DataProcessamento) And _
            Recepcao.SetConnectionAgencia(Geral.BancoCaixa) And _
            Recepcao.SetUsuario(Geral.Usuario.Login)) Then
                Screen.MousePointer = vbDefault
                Recepcao.ShowModal
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuRecRegistroOcorrencia_Click()

    Dim RegOcorrencia   As New MDI_RegOcorrencia.RegistroOcorrencia
    Dim ret_imp         As Integer
    
    Screen.MousePointer = vbHourglass

    'Inicia autenticadora
    ''''''''''''''''''''''''''''''''''''
    ' Verifica se é impressora IBM (1) '
    ' ou PROCOMP (2)                   '
    ''''''''''''''''''''''''''''''''''''
    Geral.Autenticadora = 1
    If Geral.Autenticadora = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Para Registro de Ocorrência é necessário que alguma autenticadora esteja ligada a estação. Verifique se existe alguma autenticadora ligada a esta estação e a selecione no módulo de Parâmetros do Sistema.", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    Else
        On Error GoTo ErroAutentica

        'ret_imp = Autentica.Inicia()

        On Error GoTo 0

        If (ret_imp <> 0) Then
            Screen.MousePointer = vbDefault
            MsgBox "A Autenticadora não está respondendo. Verifique se ela está ligada!", vbExclamation + vbOKOnly, App.Title
            Exit Sub
        End If

        If Not RegOcorrencia.SetConnection(Geral.Banco) Then
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível estabelecer a conexão com o Banco de Dados.", vbExclamation
        Else
            If (RegOcorrencia.SetDataProcessamento(Geral.DataProcessamento) And _
                RegOcorrencia.SetUsuario(Geral.Usuario.Login) And _
                RegOcorrencia.SetAutenticadora(Geral.Autenticadora) And _
                RegOcorrencia.SetAutentica(Autentica) And _
                RegOcorrencia.SetAgenciaCentral(Geral.AgenciaCentral) And _
                RegOcorrencia.SetAgenciaApresentante(Geral.AgenciaApresentante) And _
                RegOcorrencia.SetConnectionAgencia(Geral.BancoCaixa)) Then
                
                Screen.MousePointer = vbDefault
                RegOcorrencia.ShowModal
            End If
        End If

    End If

    If (Geral.Autenticadora <> 0) And (Not Autentica Is Nothing) Then
        Autentica.Finaliza
    End If

    Screen.MousePointer = vbDefault
    Exit Sub

ErroAutentica:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível iniciar a Autenticadora. Verifique se o arquivo .DLL da autenticadora se encontra no diretório do Windows.", vbInformation + vbOKOnly, App.Title
    
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

