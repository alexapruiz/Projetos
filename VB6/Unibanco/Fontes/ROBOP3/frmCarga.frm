VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCarga 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Tabelas"
   ClientHeight    =   3270
   ClientLeft      =   2550
   ClientTop       =   2535
   ClientWidth     =   4890
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCopia 
      Height          =   3255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4875
      Begin VB.PictureBox PictureCopia 
         Height          =   555
         Left            =   195
         Picture         =   "frmCarga.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   525
         TabIndex        =   19
         Top             =   165
         Width           =   585
      End
      Begin VB.Frame Frame5 
         Height          =   1260
         Left            =   150
         TabIndex        =   14
         Top             =   660
         Width           =   4515
         Begin MSComCtl2.Animation AnimationCopia 
            Height          =   1005
            Left            =   90
            TabIndex        =   15
            Top             =   165
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1773
            _Version        =   393216
            FullWidth       =   287
            FullHeight      =   67
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1260
         Left            =   150
         TabIndex        =   13
         Top             =   1920
         Width           =   4515
         Begin VB.Label LabelCopiandoArquivos 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   16
            Top             =   540
            Width           =   4245
         End
      End
      Begin VB.Label LabelCopia 
         AutoSize        =   -1  'True
         Caption         =   "Cópia Física de Arquivos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1305
         TabIndex        =   20
         Top             =   330
         Width           =   2685
      End
   End
   Begin VB.Frame FrameCarga 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      Begin MSComCtl2.Animation AnimationCarga 
         Height          =   555
         Left            =   255
         TabIndex        =   17
         Top             =   150
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   979
         _Version        =   393216
         FullWidth       =   39
         FullHeight      =   37
      End
      Begin VB.Frame Frame1 
         Height          =   1275
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   4515
         Begin VB.PictureBox PictureProgressTabela 
            BackColor       =   &H00808080&
            Height          =   405
            Left            =   225
            ScaleHeight     =   345
            ScaleWidth      =   3990
            TabIndex        =   8
            Top             =   444
            Width           =   4050
            Begin VB.Label LabelProgress 
               BackColor       =   &H00800000&
               Height          =   375
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   4005
            End
         End
         Begin VB.Label LabelTabelaCorrente 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Carregando a tabela Conax .... "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   915
            TabIndex        =   11
            Top             =   135
            Width           =   2700
         End
         Begin VB.Label LabelQtdeRegistros 
            AutoSize        =   -1  'True
            Caption         =   "Qtd. de registros carregados: 00000000"
            ForeColor       =   &H00000040&
            Height          =   192
            Left            =   888
            TabIndex        =   10
            Top             =   984
            Width           =   2808
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1230
         Left            =   240
         TabIndex        =   1
         Top             =   1935
         Width           =   4515
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00808080&
            Height          =   405
            Left            =   225
            ScaleHeight     =   345
            ScaleWidth      =   3990
            TabIndex        =   2
            Top             =   450
            Width           =   4050
            Begin VB.Label LabelProgressAll 
               BackColor       =   &H000000C0&
               Height          =   375
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.Label LabelTabelasCarregadas 
            AutoSize        =   -1  'True
            Caption         =   "Tabelas Carregadas: 0"
            ForeColor       =   &H00000040&
            Height          =   192
            Left            =   1068
            TabIndex        =   6
            Top             =   984
            Width           =   1608
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Status Geral"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1620
            TabIndex        =   5
            Top             =   135
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "de 10."
            ForeColor       =   &H00000040&
            Height          =   192
            Left            =   2832
            TabIndex        =   4
            Top             =   984
            Width           =   456
         End
      End
      Begin VB.Label LabelCarga 
         AutoSize        =   -1  'True
         Caption         =   "Transferindo Tabelas -> MDI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   18
         Top             =   330
         Width           =   4065
      End
   End
End
Attribute VB_Name = "frmCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GareReceita As estruturaGareReceita
Private GareGrupo   As estruturaGareGrupo

'Declaração do Type do Registro Gare Receita
Private Type estruturaGareReceita
    CodigoPagamento                 As String * 4
    CodigoGrupo                     As String * 1
    Data_Inicio_Vigencia            As String * 8
    Data_Final_Vigencia             As String * 8
    Indicador_Excessao              As String * 1
    Indicador_Arrecadacao           As String * 1
    Tipo_Servico                    As String * 1
    Indicador_Autenticacao          As String * 1
    Indicador_Servico_Autenticacao  As String * 3
    Numero_Vias_Comprovante         As String * 2
    Valor                           As String * 12
    Caracter                        As String * 1
End Type

'Declaração do Type do Registro Gare Grupo
Private Type estruturaGareGrupo
    CodigoGrupoReceita              As String * 1
    IndicadorCotaIPVA               As String * 1
    IndicadorVenctoNormal           As String * 1
    IndicadorInscEstadual           As String * 1
    
    IndicadorCampoDocto             As String * 1
    IndicadorInscAtiva              As String * 1
    IndicadorReferencia             As String * 1
    IndicadorNumParcelamento        As String * 1
    
    IndicadorValorReceita           As String * 1
    IndicadorValorJuro              As String * 1
    IndicadorValorMulta             As String * 1
    IndicadorAcresFinanceiro        As String * 1
    
    IndicadorHonoAdvogado           As String * 1
    Caracter                        As String * 1
End Type
Sub CopiaArquivos()

    Dim DiretorioTabelas    As String
    
    Screen.MousePointer = 11
    
    If Dir(App.Path & "\Util\FILEMOVE.AVI", vbArchive) <> "" Then
        AnimationCopia.Open App.Path & "\Util\FILEMOVE.AVI"
        AnimationCopia.Play
    End If
        
    LabelCopiandoArquivos.Caption = "Localizando Arquivos ..."
    Espera (0.5)
    
    DiretorioTabelas = PegarOpcaoINI("Path", "Tabelas", "1")
    
    If Dir(DiretorioTabelas & "\NVSCRGPS.DAT", vbArchive) = "" Or Dir(DiretorioTabelas & "\NVSCDARF.DAT", vbArchive) = "" Or _
       Dir(DiretorioTabelas & "\TFSCRGPS.DAT", vbArchive) = "" Or Dir(DiretorioTabelas & "\TFSCDARF.DAT", vbArchive) = "" Then
    
        If MsgBox("Os arquivos de GPS e DARF não foram localizados, Continua...", vbCritical + vbOKCancel, App.Title) = vbCancel Then
            End
        End If
    Else
        If Dir(App.Path & "\Tabelas", vbDirectory) = "" Then
            MkDir App.Path & "\Tabelas"
        End If
        If Dir(App.Path & "\Tabelas\*.*", vbArchive) <> "" Then
            Kill App.Path & "\Tabelas\*.*"
        End If
        
       'GPS
        LabelCopiandoArquivos.Caption = "Copiando Arquivo: Validação: GPS"
        Espera (0.9)
            
        FileCopy DiretorioTabelas & "\NVSCRGPS.DAT", App.Path & "\Tabelas\GPS.TXT"
        
       'DARF
        LabelCopiandoArquivos.Caption = "Copiando Arquivo: Validação: DARF"
        Espera (0.9)
        
        FileCopy DiretorioTabelas & "\NVSCDARF.DAT", App.Path & "\Tabelas\DARF.TXT"
        
       'GARE Grupo
        LabelCopiandoArquivos.Caption = "Copiando Arquivo: Validação: GARE Grupo"
        Espera (0.9)
        
        FileCopy DiretorioTabelas & "\TFSCRGPS.DAT", App.Path & "\Tabelas\GAREGRUPO.TXT "
        
       'GARE Receita
        LabelCopiandoArquivos.Caption = "Copiando Arquivo: Validação Gare Receita"
        Espera (0.9)
        
        FileCopy DiretorioTabelas & "\TFSCDARF.DAT", App.Path & "\Tabelas\GARERECEITA.TXT "
        
        LabelCopiandoArquivos.Caption = "Copias Efetuadas com Sucesso ..."
        Screen.MousePointer = 0
        Espera (1)
    End If
    
    AnimationCopia.Stop
    FrameCopia.Visible = False
    
    Exit Sub
    
TrataErro:
    Screen.MousePointer = 0
    MsgBox "Falha na copia de Arquivos. ", vbOKOnly + vbCritical, "Atenção"
    End
    Exit Sub

    
End Sub
Sub CarregarDarfGps()
   On Error GoTo TrataErro

    Dim ArqDarf             As Double
    Dim ArqGps              As Double
    Dim ArqGareGrupo        As Double
    Dim ArqGareReceita      As Double
    Dim Tam                 As Double
    Dim Cont                As Double
    Dim DetalheDarf         As String * 58
    Dim DetalheGps          As String * 39
    Dim DarfData            As Long
    Dim DarfExcDocto        As String
    Dim Passo               As Double
               
    LabelTabelaCorrente.Caption = "7. Carregando o Arquivo DARF PRETO .... "
    DoEvents

   '===========================================================================================
   '7.) Leitura do arquivo DARF_PRETO
    Tam = FileLen(App.Path & "\TABELAS\DARF.TXT") / 58
    Passo = LenLabelProgress / Tam
               
    Cont = 0
    LabelProgress.Width = 0
    DetalheDarf = ""
    
    ArqDarf = FreeFile
    Open App.Path & "\TABELAS\DARF.TXT" For Binary As #ArqDarf
    
    Do
        Cont = Cont + 1
        LabelProgress.Width = LabelProgress.Width + Passo
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
        Get ArqDarf, , DetalheDarf
        
       'gravação dos dados lidos na tabela DarfPreto
        If Val(Mid(DetalheDarf, 51, 4)) < 1900 Then
           DarfData = Val(Mid(DetalheDarf, 51, 4)) + 2000
        Else
           DarfData = Val(Mid(DetalheDarf, 51, 4))
        End If
        
        If Mid(DetalheDarf, 55, 1) = " " Then
           DarfExcDocto = "0"
        Else
           DarfExcDocto = Mid(DetalheDarf, 55, 1)
        End If
        
        If Asc(DetalheDarf) = 0 Or EOF(ArqDarf) Then
            Exit Do
        End If
        
        MDIQuery.insValidaDarfPreto DetalheDarf, DarfExcDocto, DarfData
        
    Loop Until EOF(ArqDarf)
    
    Close #ArqDarf
    
    LabelProgressAll.Width = 2800
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 07"
    LabelTabelaCorrente.Caption = "8. Carregando o Arquivo GPS .... "
    DoEvents

   '===========================================================================================
   '8.) Leitura do arquivo GPS
       
    Tam = FileLen(App.Path & "\TABELAS\GPS.TXT") / 39
    Passo = LenLabelProgress / Tam
           
    Cont = 0
    LabelProgress.Width = 0
    DetalheGps = ""
    
    ArqGps = FreeFile

    Open App.Path & "\TABELAS\GPS.TXT" For Binary As #ArqGps
    
    Do
        Cont = Cont + 1
        Espera (0.1)
        LabelProgress.Width = LabelProgress.Width + Passo
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
        Get ArqDarf, , DetalheGps

        If Mid(DetalheGps, 1, 4) = "9999" Or EOF(ArqGps) Then
            Exit Do
        End If
        
       'gravação dos dados da tabela GPS
        MDIQuery.insValidaGps DetalheGps
    Loop Until EOF(ArqGps)
    
    Close #ArqGps
    
    LabelProgressAll.Width = 3200
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 08"
    LabelTabelaCorrente.Caption = "9. Carregando o Arquivo Gare Grupo .... "
    DoEvents

    
   '===========================================================================================
   '9.) Leitura do arquivo GAREGRUPO
    Tam = FileLen(App.Path & "\TABELAS\GAREGRUPO.TXT") / 16
    Passo = LenLabelProgress / Tam
               
    Cont = 0
    LabelProgress.Width = 0
    
    ArqGareGrupo = FreeFile
    Open App.Path & "\TABELAS\GAREGRUPO.TXT" For Binary As #ArqGareGrupo
    
    Do
        Cont = Cont + 1
        Espera (0.1)
        LabelProgress.Width = LabelProgress.Width + Passo
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
        Get ArqGareGrupo, , GareGrupo
        
        If EOF(ArqGareGrupo) Then Exit Do
                
       'gravação dos dados da tabela Grupo Gare
        MDIQuery.insRegraMdi_GrupoGare GareGrupo.CodigoGrupoReceita, GareGrupo.IndicadorCotaIPVA, GareGrupo.IndicadorVenctoNormal, GareGrupo.IndicadorInscEstadual, GareGrupo.IndicadorCampoDocto, GareGrupo.IndicadorInscAtiva, GareGrupo.IndicadorReferencia, GareGrupo.IndicadorNumParcelamento, GareGrupo.IndicadorValorReceita, GareGrupo.IndicadorValorJuro, GareGrupo.IndicadorValorMulta, GareGrupo.IndicadorAcresFinanceiro, GareGrupo.IndicadorHonoAdvogado

                
    Loop
    
    Close #ArqGareGrupo
    
    LabelProgressAll.Width = 3600
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 09"
    LabelTabelaCorrente.Caption = "10. Carregando o Arquivo Gare Receita .... "
    DoEvents
    
'===========================================================================================
   '10.) Leitura do arquivo Gare Receita
    Tam = FileLen(App.Path & "\TABELAS\GARERECEITA.TXT") / 43
    Passo = LenLabelProgress / Tam
               
    Cont = 0
    LabelProgress.Width = 0
    
    ArqGareReceita = FreeFile
    Open App.Path & "\TABELAS\GARERECEITA.TXT" For Binary As #ArqGareReceita
    
    Do
        Cont = Cont + 1
        LabelProgress.Width = LabelProgress.Width + Passo
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
        Get ArqGareReceita, , GareReceita
        If EOF(ArqGareReceita) Then Exit Do
                
        MDIQuery.insRegraMdi_Gare GareReceita.CodigoPagamento, GareReceita.CodigoGrupo, GareReceita.Data_Inicio_Vigencia, GareReceita.Data_Final_Vigencia, GareReceita.Indicador_Excessao, GareReceita.Indicador_Arrecadacao, GareReceita.Tipo_Servico, GareReceita.Indicador_Autenticacao, GareReceita.Indicador_Servico_Autenticacao, GareReceita.Numero_Vias_Comprovante, GareReceita.Valor
        
    Loop
    
    Close #ArqGareReceita
    
    LabelProgressAll.Width = 4005
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 10"
 
    Espera (2)
    
    Exit Sub
    
TrataErro:
    Screen.MousePointer = 0
    MsgBox "Não foi possivel carregar as tabelas. Reinicialize e tente novamente", vbOKOnly + vbCritical, "Atenção"
    End
    Exit Sub
    
'   Resume

End Sub
Sub CargaGeral()

    Dim Cont                As Double
    Dim Passo               As Double
    Dim Ret                 As Boolean
    Dim RstMDI              As Recordset
    Dim RstUBB              As Recordset
    
    Me.Refresh
    Screen.MousePointer = 11
    
    If Dir(App.Path & "\util\down5.avi", vbArchive) <> "" Then
        AnimationCarga.Open App.Path & "\util\down5.avi"
        AnimationCarga.Play
    End If

    LabelTabelaCorrente.Caption = "Deletando as Tabelas antigas .... "
    DoEvents

   'exclusão das tabelas
    LabelProgress.BackColor = QBColor(4)
    LabelProgress.Width = 4005
    Espera (0.2)
    
    LabelProgress.Width = 3670
    MDIQuery.delTabelas "conax"
    Espera (0.2)
       
    LabelProgress.Width = 3340
    MDIQuery.delTabelas "abgag"
    Espera (0.2)
    
    LabelProgress.Width = 3000
    MDIQuery.delTabelas "agenf"
    Espera (0.2)
    
    LabelProgress.Width = 2650
    MDIQuery.delTabelas "ValidaDarfPreto"
    Espera (0.2)
    
    LabelProgress.Width = 2300
    MDIQuery.delTabelas "ValidaGARE"
    Espera (0.2)
    
    LabelProgress.Width = 1950
    MDIQuery.delTabelas "ValidaGPS"
    Espera (0.2)
    
    LabelProgress.Width = 1500
    MDIQuery.delTabelas "TfsBanco"
    Espera (0.2)
    
    LabelProgress.Width = 1100
    MDIQuery.delTabelas "TfsCCred"
    Espera (0.2)
    
    LabelProgress.Width = 750
    MDIQuery.delTabelas "RegraMdi_GrupoGare"
    Espera (0.2)
    
    LabelProgress.Width = 380
    MDIQuery.delTabelas "RegraMdi_Gare"
    Espera (0.2)
    
    LabelProgress.BackColor = QBColor(1)
    LabelProgress.Width = 0
    Espera (0.2)
    
   'Carga das 10 tabelas do UBB-NT - TfsConax, TfsAgeng, TfsAbgag, TfsDarfPreto, TfsGare, TfsGps, TfsBanco, TfsCCred, RegraMDI_Gare, RegraMDI_GrupoGare
    LabelTabelaCorrente.Caption = "1. Carregando a Tabela CONAX .... "
    DoEvents

   '===========================================================================================
   ' 1.) Leitura da tabela CONAX do UBB-NT
   
    Set RstUBB = UBBQuery.getTabela("tfsConax")
    
   'Inicia label progress
    LabelProgress.Width = 0
       
   'Se não encontrar nenhuma linha na tabela tfsconax
    If RstUBB.EOF() Then
       Screen.MousePointer = 0
       MsgBox "Atenção! A tabela de arrecadação Conax está vazia. ", vbOKOnly + vbCritical, "Atenção"
       End
       Exit Sub
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
        
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
       
       'gravação dos dados da tabela conax
        MDIQuery.insConax RstUBB!confscdprod, RstUBB!confscdfebr, RstUBB!confscdsegu, _
                                              RstUBB!confsaxprod, RstUBB!confstparre, _
                                              RstUBB!confsstrepa, RstUBB!confsnoprod, _
                                              RstUBB!confsstbarr
             
        RstUBB.MoveNext
        LabelProgress.Width = Cont * Passo
        
    Loop Until RstUBB.EOF
    
    RstUBB.Close
    
    LabelProgressAll.Width = 400
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 01"
    LabelTabelaCorrente.Caption = "2. Carregando a Tabela ABGAG .... "
    DoEvents
   '===========================================================================================
   ' 2.) Leitura da tabela ABGAG do UBB-NT
    Set RstUBB = UBBQuery.getTabela("tfsabgag")
    
   'Inicia label progress
    LabelProgress.Width = 0
        
   'se nao encontrar nenhuma linha na tabela tfsabgag
    If RstUBB.EOF() Then
        Screen.MousePointer = 0
        MsgBox "Atenção! A tabela de arrecadação Abgag está vazia. ", vbOKOnly + vbCritical, "Atenção"
        End
        Exit Sub
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
    
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
       'gravação dos dados da tabela abgag
        MDIQuery.insABGAG RstUBB!abgfscdprod, RstUBB!abgfsabagen, RstUBB!abgfscdagen
                
        LabelProgress.Width = Cont * Passo
        RstUBB.MoveNext
        
   Loop Until RstUBB.EOF
   
   LabelProgressAll.Width = 800
   LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 02"
   LabelTabelaCorrente.Caption = "3. Carregando a Tabela AGENF .... "
   DoEvents

   '===========================================================================================
   '3.) Leitura da tabela AGENG do UBB-NT
    Set RstUBB = UBBQuery.getTabelaAgeng
    
   'Inicia label progress
    LabelProgress.Width = 0
       
    If RstUBB.EOF() Then
        Screen.MousePointer = 0
        MsgBox "Atenção! A tabela de arrecadação Agenf está vazia. ", vbOKOnly + vbCritical, "Atenção"
        RstUBB.Close
        End
        Exit Sub
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
    
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
       'gravação dos dados da tabela abgag
        MDIQuery.insAgenf RstUBB!agefsnoagen, RstUBB!agefsestado, RstUBB!agefscdagen, _
                          RstUBB!agefsstmovi, RstUBB!agefsdtmvan, RstUBB!agefsdtmvat, _
                          RstUBB!agefsdtprox
        
        LabelProgress.Width = Cont * Passo
        RstUBB.MoveNext
    Loop Until RstUBB.EOF
    
    LabelProgressAll.Width = 1200
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 03"
    LabelTabelaCorrente.Caption = "4. Carregando a Tabela GARE .... "
    DoEvents

   '===========================================================================================
   '4.) Leitura da tabela GARE do UBB-NT
   
    Set RstUBB = UBBQuery.getTabela("tfscdrec")
    
   'Inicia label progress
    LabelProgress.Width = 0
     
   'se nao encontrar nenhuma linha na tabela tfscrec
    If RstUBB.EOF() Then
        Screen.MousePointer = 0
        MsgBox "Atenção! A tabela de Gare está vazia. ", vbOKOnly + vbCritical, "Atenção"
        RstUBB.Close
        End
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
        
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
       'gravação dos dados da tabela Gare
        MDIQuery.insValidaGare RstUBB!recfscodigo, RstUBB!recfscgrupo
        
        RstUBB.MoveNext
        LabelProgress.Width = Cont * Passo
        
    Loop Until RstUBB.EOF

    LabelProgressAll.Width = 1600
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 04"
    LabelTabelaCorrente.Caption = "5. Carregando a Tabela TFSBANCO .... "
    DoEvents
    
   '===========================================================================================
   '5.) Leitura da tabela TFSBANCO do UBB-NT
   
    Set RstUBB = UBBQuery.getTabela("tfsBanco")
    
   'Inicia label progress
    LabelProgress.Width = 0
    
   'se nao encontrar nenhuma linha na tabela tfsbanco
    If RstUBB.EOF() Then
        Screen.MousePointer = 0
        MsgBox "Atenção! A tabela de bancos validos tfsbanco está vazia. Contate suporte técnico .", vbOKOnly + vbCritical, "Atenção"
        RstUBB.Close
        End
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
    
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents
        
       'gravação dos dados da tabela tfsbanco
        MDIQuery.insTfsBanco RstUBB!banfscdbanc, RstUBB!banfsnobanc
                
        RstUBB.MoveNext
        LabelProgress.Width = Cont * Passo
        
    Loop Until RstUBB.EOF
    
   'Inicia label progress
    LabelProgressAll.Width = 2000
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 05"
    LabelTabelaCorrente.Caption = "6. Carregando a Tabela TFSCCRED .... "
    DoEvents

   '===========================================================================================
   '6.) Leitura da tabela TFSCCRED do UBB-NT
      
    Set RstUBB = UBBQuery.getTabela("tfsCcred")
    
   'Inicia label progress
    LabelProgress.Width = 0
 
   'se nao encontrar nenhuma linha na tabela tfsbanco
    If RstUBB.EOF() Then
        Screen.MousePointer = 0
        MsgBox "Atenção! A tabela TFSCCRED está vazia. Contate suporte técnico .", vbOKOnly + vbCritical, "Atenção"
        RstUBB.Close            'fecha tabela tfsbanco
        End
    End If
   
    Cont = 0
    Passo = LenLabelProgress / RstUBB.RecordCount
    
    Do
        Cont = Cont + 1
        LabelQtdeRegistros.Caption = "Qtde de registros carregados: " & Format(Cont, "00000000")
        DoEvents

       'gravação dos dados da tabela tfsccred
        MDIQuery.insTfsCred RstUBB!crefsnubinc, RstUBB!crefscdbunc, RstUBB!crefsdescri, _
                            RstUBB!crefscdsaqu, RstUBB!crefscdagen, RstUBB!crefsnuccor, _
                            RstUBB!crefsbandei
        
        RstUBB.MoveNext
        LabelProgress.Width = Cont * Passo
        
    Loop Until RstUBB.EOF
    
    LabelProgressAll.Width = 2400
    LabelTabelasCarregadas.Caption = "Tabelas Carregadas: 06"
   
   '-----------------------------------------------------------
   ' Carregar os dados de DARF_PRETO e GPS de arquivos textos '
   '-----------------------------------------------------------
       
    CarregarDarfGps
   
   'atualizar tabela parametro como tabelas UBB-NT já carregadas
    MDIQuery.updCargaTabelas Geral.DataProcessamento
    
   'exibe mensagem de final de carga
    LabelTabelasCarregadas.Caption = "Término da carga das tabelas."
    DoEvents
    
    PictureProgressTabela.Visible = False
    
    AnimationCarga.Stop
    Unload Me
    
    Exit Sub

TrataErro:
    Screen.MousePointer = 0
    MsgBox "Não foi possivel carregar as tabelas. Reinicialize e tente novamente", vbOKOnly + vbCritical, "Atenção"
    End
    Exit Sub

End Sub
Private Sub Form_Activate()
    Call CopiaArquivos
    Call CargaGeral
End Sub
