VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Importacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de Dados e Imagens"
   ClientHeight    =   3960
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4104
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   396
      Left            =   2094
      TabIndex        =   10
      Top             =   3504
      Width           =   1356
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar"
      Height          =   396
      Left            =   654
      TabIndex        =   9
      Top             =   3504
      Width           =   1356
   End
   Begin VB.Frame Frame3 
      Height          =   636
      Left            =   96
      TabIndex        =   8
      Top             =   2784
      Width           =   3900
      Begin ComctlLib.ProgressBar gauImportacao 
         Height          =   252
         Left            =   96
         TabIndex        =   13
         Top             =   240
         Width           =   3708
         _ExtentX        =   6541
         _ExtentY        =   445
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1980
      Left            =   96
      TabIndex        =   1
      Top             =   816
      Width           =   3900
      Begin VB.Label lblAgencia 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   2112
         TabIndex        =   12
         Top             =   288
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Agência:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   576
         TabIndex        =   11
         Top             =   288
         Width           =   780
      End
      Begin VB.Label lblDoctos 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   2112
         TabIndex        =   7
         Top             =   1488
         Width           =   1212
      End
      Begin VB.Label lblCapas 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   2112
         TabIndex        =   6
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label lblLotes 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   2112
         TabIndex        =   5
         Top             =   672
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Documentos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   576
         TabIndex        =   4
         Top             =   1512
         Width           =   1164
      End
      Begin VB.Label Label2 
         Caption         =   "Capas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   576
         TabIndex        =   3
         Top             =   1080
         Width           =   684
      End
      Begin VB.Label Label1 
         Caption         =   "Lotes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   576
         TabIndex        =   2
         Top             =   696
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Localizar Drive"
      Height          =   732
      Left            =   96
      TabIndex        =   0
      Top             =   48
      Width           =   3900
      Begin VB.ComboBox cmbDrive 
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
         Height          =   336
         Left            =   144
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   3612
      End
   End
End
Attribute VB_Name = "Importacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDriveType Lib "kernel32" _
    Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2

Private Const FO_COPY = &H2
Private Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Private Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Private Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files

Private qryVerificaImportacao       As rdoQuery
Private qryInsereAgencia            As rdoQuery
Private qryInsereLote               As rdoQuery
Private qryInsereCapa               As rdoQuery
Private qryInsereDocumento          As rdoQuery
Private qryInsereLog                As rdoQuery
Private qryInsereControleImportacao As rdoQuery

Private m_Agencia                   As Integer
Private m_Remessa                   As Long
Private m_CD                        As Byte
Private m_Lotes                     As Long
Private m_Capas                     As Long
Private m_Doctos                    As Long
Private m_Horas                     As String

Private Type SHFILEOPSTRUCT
        hwnd                        As Long
        wFunc                       As Long
        pFrom                       As String
        pTo                         As String
        fFlags                      As Integer
        fAnyOperationsAborted       As Long
        hNameMappings               As Long
        lpszProgressTitle           As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Sub Preenche_cmbDrive()
    Dim i               As Integer
    Dim StrDrive        As String
    Dim DriveType       As Long
    
    cmbDrive.Clear
    For i = 67 To 90
        StrDrive = Chr(i) & ":"
        DriveType = GetDriveType(StrDrive)
        If DriveType = DRIVE_CDROM Or DriveType = DRIVE_REMOVABLE Then
            cmbDrive.AddItem StrDrive
        End If
    Next
End Sub

Private Function ValidaCD_ID(ByVal Drive As String) As Boolean
    Dim StrFile         As String
    Dim StrReg          As String
    Dim CDFile          As Integer
    Dim DataProc        As Long
    
    ValidaCD_ID = False
    On Error GoTo ErroFile
    
    StrFile = Drive & "\" & DIR_DADOS & Geral.DataProcessamento & "\" & ARQ_CD
    If Dir(StrFile, vbArchive) = ARQ_CD Then
        CDFile = FreeFile
        Open StrFile For Input Access Read Lock Read Write As #CDFile
        Input #CDFile, StrReg
        Close #CDFile
        If Len(StrReg) = 27 Then
            m_Agencia = CInt(Left(StrReg, 4))
            DataProc = CLng(Mid(StrReg, 5, 8))
            m_Horas = Mid(StrReg, 13, 8)
            m_Remessa = CLng(Mid(StrReg, 21, 5))
            m_CD = CByte(Right(StrReg, 2))
            If m_Agencia = 0 Then
                MsgBox "A Agência não é valida.", vbOKOnly + vbCritical, App.Title
                Exit Function
            End If
            If DataProc <> Geral.DataProcessamento Then
                MsgBox "A Data do CD não confere com a Data do Movimento.", vbOKOnly + vbCritical, App.Title
                Exit Function
            End If
            If m_Remessa = 0 Then
                MsgBox "o Número da Remessa não é valido.", vbOKOnly + vbCritical, App.Title
                Exit Function
            End If
            If m_CD = 0 Then
                MsgBox "o Número do CD não é valido.", vbOKOnly + vbCritical, App.Title
                Exit Function
            End If
            ' verificar na base se este CD ainda nao foi lido
            qryVerificaImportacao.rdoParameters(0).Direction = rdParamReturnValue
            qryVerificaImportacao.rdoParameters(1) = Geral.DataProcessamento
            qryVerificaImportacao.rdoParameters(2) = m_Agencia
            qryVerificaImportacao.rdoParameters(3) = m_Remessa
            qryVerificaImportacao.rdoParameters(4) = m_CD
            qryVerificaImportacao.Execute
            If qryVerificaImportacao.rdoParameters(0) = 0 Then
                ValidaCD_ID = True
            Else
                MsgBox "Este CD já foi Importado anteriormente.", vbOKOnly + vbExclamation, App.Title
            End If
        Else
            MsgBox "O Arquivo " & ARQ_CD & " não está no formato esperado.", vbOKOnly + vbCritical, App.Title
        End If
    Else
        MsgBox "Não foi possível localizar o arquivo " & StrFile, vbOKOnly + vbCritical, App.Title
    End If
    Exit Function
    
ErroFile:
    TratamentoErro "Erro na leitura do Header do CD", Err, rdoErrors, True
End Function

Private Function CopiarImagens(ByVal PathOrigem As String, ByVal PathDestino As String) As Boolean
    Dim FHStruct        As SHFILEOPSTRUCT
    Dim Result          As Long
    
    rdoErrors.Clear
    On Error GoTo ErroCopia
    
    FHStruct.wFunc = FO_COPY
    FHStruct.pFrom = PathOrigem
    FHStruct.pTo = PathDestino
    FHStruct.fFlags = FOF_SIMPLEPROGRESS + FOF_NOCONFIRMATION + FOF_NOCONFIRMMKDIR
    
    Result = SHFileOperation(FHStruct)
    
    If Result = 0 Then
        CopiarImagens = True
    Else
        CopiarImagens = False
        MsgBox "Erro na cópia das Imagens. Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    End If
    Exit Function
    
ErroCopia:
    TratamentoErro "Erro na cópia das Imagens. Importação NÃO foi completada.", Err, rdoErrors, True
End Function

Private Function ImportarAgencia(ByVal PathName As String) As Boolean
    Dim AgFile      As Integer
    Dim Seq         As Long
    Dim Agencia     As cg_AGENCIA
    Dim Reg         As String * 1
    Dim OffSet      As Long
    
    ImportarAgencia = False
    rdoErrors.Clear
    On Error GoTo ErroLeitura
    
    Seq = 1
    AgFile = FreeFile
    Open PathName For Binary Access Read Lock Read Write As #AgFile
    
    OffSet = 1
    While Not EOF(AgFile)
        Get #AgFile, OffSet, Reg
        Select Case Reg
            Case "H":
                Get #AgFile, OffSet, Agencia.Header
                If Seq <> CLng(Agencia.Header.Sequencial) Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Agencia.Header.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                If Not ValidaHeader(Agencia.Header, PathName) Then
                    GoTo FimLeitura
                End If
                OffSet = OffSet + Len(Agencia.Header)
            Case "R":
                Get #AgFile, OffSet, Agencia.Registro
                If Agencia.Registro.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Agencia.Registro.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                qryInsereAgencia.rdoParameters(0).Direction = rdParamReturnValue
                qryInsereAgencia.rdoParameters(1) = Geral.DataProcessamento
                qryInsereAgencia.rdoParameters(2) = CInt(Agencia.Registro.Agencia)
                qryInsereAgencia.rdoParameters(3) = CLng(Agencia.Registro.Lacre)
                qryInsereAgencia.rdoParameters(4) = CInt(Agencia.Registro.QtdInformada)
                qryInsereAgencia.rdoParameters(5) = Agencia.Registro.HoraCadastrada
                qryInsereAgencia.rdoParameters(6) = Agencia.Registro.IdEnv_Mal
                qryInsereAgencia.Execute
                If qryInsereAgencia.rdoParameters(0) <> 0 Then
                    MsgBox "Erro na importação dos dados da agência. Importação NÃO foi completada.", _
                        vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                OffSet = OffSet + Len(Agencia.Registro)
            Case "T":
                Get #AgFile, OffSet, Agencia.Trailler
                If Agencia.Trailler.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Agencia.Trailler.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                If Not ValidaTrailler(Agencia.Trailler, PathName) Then
                    GoTo FimLeitura
                End If
                OffSet = OffSet + Len(Agencia.Trailler)
        End Select
        
        Seq = Seq + 1
    Wend
    
    Close #AgFile
    ImportarAgencia = True
    Exit Function
    
FimLeitura:
    Close #AgFile
    Exit Function
    
ErroLeitura:
    TratamentoErro "Erro na importação dos dados da agência. Importação NÃO foi completada.", Err, rdoErrors, True
End Function

Private Function ImportarDados(ByVal PathName As String) As Boolean
    Dim DatFile     As Integer
    Dim Seq         As Long
    Dim Dados       As cg_DADOS
    Dim Reg         As String * 1
    Dim OffSet      As Long
    Dim IdLote      As Long
    Dim IdCapa      As Long
    Dim IdDocto     As Long
    
    ImportarDados = False
    rdoErrors.Clear
    On Error GoTo ErroLeitura
    
    Seq = 1
    DatFile = FreeFile
    Open PathName For Binary Access Read Lock Read Write As #DatFile
    
    m_Lotes = 0
    m_Capas = 0
    m_Doctos = 0
    
    IdLote = 0
    IdCapa = 0
    IdDocto = 0
    
    OffSet = 1
    
    While Not EOF(DatFile)
        Get #DatFile, OffSet, Reg
        Select Case Reg
            Case "H":
                Get #DatFile, OffSet, Dados.Header
                OffSet = OffSet + Len(Dados.Header)
                
                If Seq <> CLng(Dados.Header.Sequencial) Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Header.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                If Not ValidaHeader(Dados.Header, PathName) Then
                    GoTo FimLeitura
                End If
            
            Case "L": 'Lote
                Get #DatFile, OffSet, Dados.Lote
                OffSet = OffSet + Len(Dados.Lote)
                
                If Dados.Lote.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Lote.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                If CLng(Dados.Lote.IdLote) > 0 Then
                    qryInsereLote.rdoParameters(0).Direction = rdParamReturnValue
                    qryInsereLote.rdoParameters(1) = Geral.DataProcessamento
                    qryInsereLote.rdoParameters(2) = CInt(Dados.Lote.Prioridade)
                    qryInsereLote.rdoParameters(3) = m_Agencia
                    qryInsereLote.rdoParameters(4) = CLng(Dados.Lote.IdLote)
                    
                    qryInsereLote.Execute
                    If qryInsereLote.rdoParameters(0) <> 0 Then
                        MsgBox "Erro na importação dos dados. Importação NÃO foi completada.", _
                            vbOKOnly + vbCritical, App.Title
                        GoTo FimLeitura
                    End If
                End If
                
                IdLote = Dados.Lote.IdLote
                m_Lotes = m_Lotes + 1
                
            Case "C": 'Capa
                IdDocto = 0
                
                Get #DatFile, OffSet, Dados.Capa
                OffSet = OffSet + Len(Dados.Capa)
                
                If Dados.Capa.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Capa.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
                qryInsereCapa.rdoParameters(10).Direction = rdParamOutput
                qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
                qryInsereCapa.rdoParameters(2) = IdLote
                qryInsereCapa.rdoParameters(3) = Dados.Capa.IdEnv_Mal
                qryInsereCapa.rdoParameters(4) = Dados.Capa.Capa
                qryInsereCapa.rdoParameters(5) = Dados.Capa.Num_Malote
                qryInsereCapa.rdoParameters(6) = m_Agencia
                qryInsereCapa.rdoParameters(7) = Dados.Capa.Status
                qryInsereCapa.rdoParameters(8) = Dados.Capa.Ocorrencia
                qryInsereCapa.rdoParameters(9) = CInt(Dados.Capa.Duplicidade)
                qryInsereCapa.Execute
                If qryInsereCapa.rdoParameters(0) <> 0 Then
                    MsgBox "Erro na importação dos dados. Importação NÃO foi completada.", _
                        vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                IdCapa = qryInsereCapa.rdoParameters(10)
                m_Capas = m_Capas + 1
                
            Case "D": 'Documento
                If IdCapa = 0 Then
                    MsgBox "Arquivo de Dados não está no formato certo." & vbCrLf & _
                        "Foi encontrada um Documento sem existir uma Capa." & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                Get #DatFile, OffSet, Dados.Docto
                OffSet = OffSet + Len(Dados.Docto)
                
                If Dados.Docto.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Docto.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                qryInsereDocumento.rdoParameters(0).Direction = rdParamReturnValue
                qryInsereDocumento.rdoParameters(10).Direction = rdParamOutput
                qryInsereDocumento.rdoParameters(1) = Geral.DataProcessamento
                qryInsereDocumento.rdoParameters(2) = IdCapa
                qryInsereDocumento.rdoParameters(3) = CInt(Dados.Docto.OrdemCaptura)
                qryInsereDocumento.rdoParameters(4) = CInt(Dados.Docto.TipoDocto)
                qryInsereDocumento.rdoParameters(5) = Dados.Docto.Leitura
                qryInsereDocumento.rdoParameters(6) = Dados.Docto.Frente
                qryInsereDocumento.rdoParameters(7) = Dados.Docto.Verso
                qryInsereDocumento.rdoParameters(8) = Dados.Docto.Status
                qryInsereDocumento.rdoParameters(9) = Dados.Docto.Ordem
                qryInsereDocumento.Execute
                If qryInsereDocumento.rdoParameters(0) <> 0 Then
                    MsgBox "Erro na importação dos dados. Importação NÃO foi completada.", _
                        vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                IdDocto = qryInsereDocumento.rdoParameters(10)
                m_Doctos = m_Doctos + 1
                
            Case "G": 'loG
                If IdCapa = 0 And IdDocto = 0 Then
                    MsgBox "Arquivo de Dados não está no formato certo." & vbCrLf & _
                        "Foi encontrada um Log sem existir uma Capa ou um Documento." & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                Get #DatFile, OffSet, Dados.Log
                OffSet = OffSet + Len(Dados.Log)
                
                If Dados.Log.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Log.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
                qryInsereLog.rdoParameters(0).Direction = rdParamReturnValue
                qryInsereLog.rdoParameters(1) = Geral.DataProcessamento
                qryInsereLog.rdoParameters(2) = IdCapa
                qryInsereLog.rdoParameters(3) = IdDocto
                qryInsereLog.rdoParameters(4) = Dados.Log.Data
                qryInsereLog.rdoParameters(5) = Dados.Log.Login
                qryInsereLog.rdoParameters(6) = CByte(Dados.Log.Acao)
                qryInsereLog.Execute
                If qryInsereLog.rdoParameters(0) <> 0 Then
                    MsgBox "Erro na importação dos dados. Importação NÃO foi completada.", _
                        vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                
            Case "T":
                Get #DatFile, OffSet, Dados.Trailler
                OffSet = OffSet + Len(Dados.Trailler)
                
                If Dados.Trailler.Sequencial <> Seq Then
                    MsgBox "Sequencial do registro não confere: " & _
                        CStr(Dados.Trailler.Sequencial) & vbCrLf & _
                        "Arquivo corrompido: " & PathName & vbCrLf & _
                        "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
                    GoTo FimLeitura
                End If
                If Not ValidaTrailler(Dados.Trailler, PathName) Then
                    GoTo FimLeitura
                End If
        End Select
        
        Seq = Seq + 1
    Wend
    
    Close #DatFile
    ImportarDados = True
    Exit Function
    
FimLeitura:
    Close #DatFile
    Exit Function
    
ErroLeitura:
    TratamentoErro "Erro na importação dos dados. Importação NÃO foi completada.", Err, rdoErrors, True
End Function

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdImportar_Click()
    Dim PathOrigem As String
    Dim PathDestino As String
    
    gauImportacao.Value = 0
    
    If ValidaCD_ID(cmbDrive.List(cmbDrive.ListIndex)) Then
        PathOrigem = cmbDrive.List(cmbDrive.ListIndex) & _
                     "\" & DIR_IMAGENS & CStr(Geral.DataProcessamento) & "\*.*"
        PathDestino = Left(Geral.DiretorioImagens, Len(Geral.DiretorioImagens) - 1)
        
        gauImportacao.Value = 5
        Geral.Banco.BeginTrans
        
        If ImportarAgencia(cmbDrive.List(cmbDrive.ListIndex) & "\" & _
                    DIR_DADOS & _
                    Geral.DataProcessamento & "\" & _
                    ARQ_AGENCIA) Then
            gauImportacao.Value = 35
            If ImportarDados(cmbDrive.List(cmbDrive.ListIndex) & "\" & _
                    DIR_DADOS & _
                    Geral.DataProcessamento & "\" & _
                    ARQ_DADOS) Then
                gauImportacao.Value = 70
                If InsereControleImportacao Then
                    gauImportacao.Value = 75
                    Do While Not CopiarImagens(PathOrigem, PathDestino)
                        If MsgBox("Houve um erro na cópia das Imagens." & vbCrLf & _
                            "Deseja tentar novamente?", vbYesNo + vbQuestion, App.Title) = vbNo Then
                            GoTo Erro
                        End If
                    Loop
                    
                    gauImportacao.Value = 100
                    Geral.Banco.CommitTrans
                    
                    lblAgencia.Caption = Format(m_Agencia, "0000")
                    lblLotes.Caption = CStr(m_Lotes)
                    lblCapas.Caption = CStr(m_Capas)
                    lblDoctos.Caption = CStr(m_Doctos)
                    
                    MsgBox "Importação concluída com sucesso.", vbOKOnly + vbInformation, App.Title
                Else
                    GoTo Erro
                End If
            Else
                GoTo Erro
            End If
        Else
            GoTo Erro
        End If
    End If
    gauImportacao.Value = 0
    Exit Sub
Erro:
    gauImportacao.Value = 0
    Geral.Banco.RollbackTrans

End Sub

Private Sub Form_Activate()
    gauImportacao.Value = 0
    Preenche_cmbDrive
    If cmbDrive.ListCount > 0 Then
        cmbDrive.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
    Set qryVerificaImportacao = Geral.Banco.CreateQuery("", "{? = call MDIAG_VerificaImportacao (?,?,?,?)}")
    Set qryInsereAgencia = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereAgencia( ?,?,?,?,?,? ) }")
    Set qryInsereLote = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereLote ( ?,?,?,? )}")
    Set qryInsereCapa = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereCapa ( ?,?,?,?,?,?,?,?,?,? )}")
    Set qryInsereDocumento = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereDocumento ( ?,?,?,?,?,?,?,?,?,? )}")
    Set qryInsereLog = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereLog ( ?,?,?,?,?,? )}")
    Set qryInsereControleImportacao = Geral.Banco.CreateQuery("", "{? = call MDIAG_InsereControleImportacao ( ?,?,?,?,?,?,?,? )}")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    qryVerificaImportacao.Close
    qryInsereAgencia.Close
    qryInsereLote.Close
    qryInsereCapa.Close
    qryInsereDocumento.Close
    qryInsereLog.Close
    qryInsereControleImportacao.Close
End Sub

Private Function ValidaHeader(Header As cg_Header, ByVal Arq As String) As Boolean
    ValidaHeader = False
    If m_Agencia <> CInt(Header.AgOrig) Then
        MsgBox "Agência do arquivo " & Arq & _
            " não confere com Agência do arquivo " & ARQ_CD & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    ElseIf Geral.DataProcessamento <> CLng(Header.DataProcessamento) Then
        MsgBox "Data do arquivo " & Arq & _
            " não confere com Data do Movimento. " & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    ElseIf m_Remessa <> CInt(Header.Remessa) Then
        MsgBox "Remessa do arquivo " & Arq & _
            " não confere com Remessa do arquivo " & ARQ_CD & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    Else
        ValidaHeader = True
    End If
End Function

Private Function ValidaTrailler(Trailler As cg_Trailler, ByVal Arq As String) As Boolean
    ValidaTrailler = False
    If m_Agencia <> CInt(Trailler.AgOrig) Then
        MsgBox "Agência do arquivo " & Arq & _
            " não confere com Agência do arquivo " & ARQ_CD & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    ElseIf Geral.DataProcessamento <> CLng(Trailler.DataProcessamento) Then
        MsgBox "Data do arquivo " & Arq & _
            " não confere com Data do Movimento. " & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    ElseIf m_Remessa <> CInt(Trailler.Remessa) Then
        MsgBox "Remessa do arquivo " & Arq & _
            " não confere com Remessa do arquivo " & ARQ_CD & vbCrLf & _
            "Importação NÃO foi completada.", vbOKOnly + vbCritical, App.Title
    Else
        ValidaTrailler = True
    End If
End Function

Private Function InsereControleImportacao() As Boolean

    InsereControleImportacao = False
    On Error GoTo ErroControle
    
    qryInsereControleImportacao.rdoParameters(0).Direction = rdParamReturnValue
    qryInsereControleImportacao.rdoParameters(1) = Geral.DataProcessamento
    qryInsereControleImportacao.rdoParameters(2) = CStr(Geral.DataProcessamento) & " " & m_Horas
    qryInsereControleImportacao.rdoParameters(3) = m_Agencia
    qryInsereControleImportacao.rdoParameters(4) = m_Remessa
    qryInsereControleImportacao.rdoParameters(5) = m_CD
    qryInsereControleImportacao.rdoParameters(6) = m_Lotes
    qryInsereControleImportacao.rdoParameters(7) = m_Capas
    qryInsereControleImportacao.rdoParameters(8) = m_Doctos
    qryInsereControleImportacao.Execute
    If qryInsereControleImportacao.rdoParameters(0) <> 0 Then
        MsgBox "Erro na importação dos dados. Importação NÃO foi completada.", _
            vbOKOnly + vbCritical, App.Title
        Exit Function
    End If
    InsereControleImportacao = True
    Exit Function
    
ErroControle:
    TratamentoErro "Erro na importação dos dados. Importação NÃO foi completada.", Err, rdoErrors, True
End Function
