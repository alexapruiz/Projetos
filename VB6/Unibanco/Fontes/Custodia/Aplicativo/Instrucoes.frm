VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Instrucoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instruções enviadas pelo VC."
   ClientHeight    =   6585
   ClientLeft      =   795
   ClientTop       =   1470
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   570
      Top             =   5850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   720
      Left            =   9660
      Picture         =   "Instrucoes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5730
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.CommandButton cmdNovaSelecao 
      Caption         =   "&Nova seleção"
      Height          =   720
      Left            =   9648
      Picture         =   "Instrucoes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5736
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   720
      Left            =   8784
      Picture         =   "Instrucoes.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5736
      Width           =   816
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   720
      Left            =   10536
      Picture         =   "Instrucoes.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5736
      Width           =   816
   End
   Begin TabDlg.SSTab TabOpcoes 
      Height          =   5460
      Left            =   96
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   144
      Width           =   11268
      _ExtentX        =   19897
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Argumentos de Pesquisa"
      TabPicture(0)   =   "Instrucoes.frx":0F88
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGeral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Instruções"
      TabPicture(1)   =   "Instrucoes.frx":0FA4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdInstrucoes"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraGeral 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seleção"
         ForeColor       =   &H80000008&
         Height          =   4620
         Left            =   912
         TabIndex        =   5
         Top             =   456
         Width           =   9372
         Begin VB.Frame frabordero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Identificação do Borderô"
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   744
            TabIndex        =   26
            Top             =   624
            Width           =   3948
            Begin VB.TextBox txtnumBordero 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1320
               MaxLength       =   19
               TabIndex        =   28
               Top             =   384
               Width           =   2484
            End
            Begin VB.OptionButton optbordero 
               BackColor       =   &H80000000&
               Caption         =   "Borderô"
               Height          =   348
               Left            =   300
               TabIndex        =   27
               Top             =   336
               Width           =   876
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Identificação"
               Height          =   252
               Left            =   1488
               TabIndex        =   29
               Top             =   192
               Width           =   2172
            End
         End
         Begin VB.Frame fraDeposito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Data Anterior"
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   744
            TabIndex        =   22
            Top             =   1896
            Width           =   3948
            Begin VB.OptionButton optdteAnterior 
               BackColor       =   &H80000000&
               Caption         =   "Anterior"
               Height          =   348
               Left            =   300
               TabIndex        =   23
               Top             =   336
               Width           =   1116
            End
            Begin DATEEDITLib.DateEdit DteAnterior 
               Height          =   348
               Left            =   1968
               TabIndex        =   24
               Top             =   384
               Width           =   1524
               _Version        =   65537
               _ExtentX        =   2688
               _ExtentY        =   614
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.84
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "Data"
               Height          =   252
               Left            =   2016
               TabIndex        =   25
               Top             =   192
               Width           =   1428
            End
         End
         Begin VB.Frame fracheque 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Identificação do Cheque"
            ForeColor       =   &H80000008&
            Height          =   3372
            Left            =   5388
            TabIndex        =   10
            Top             =   624
            Width           =   3180
            Begin VB.TextBox TxtCcc 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1632
               MaxLength       =   7
               TabIndex        =   19
               Top             =   1632
               Width           =   1092
            End
            Begin VB.TextBox txtcAgencia 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2208
               MaxLength       =   4
               TabIndex        =   18
               Top             =   1032
               Width           =   516
            End
            Begin VB.TextBox TxtcBco 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   17
               Top             =   432
               Width           =   444
            End
            Begin VB.OptionButton optcheque 
               BackColor       =   &H80000000&
               Caption         =   "Cheque"
               Height          =   348
               Left            =   300
               TabIndex        =   11
               Top             =   312
               Width           =   885
            End
            Begin VB.TextBox txtNumCheque 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1632
               MaxLength       =   7
               TabIndex        =   20
               Top             =   2208
               Width           =   1092
            End
            Begin CURRENCYEDITLib.CurrencyEdit txtValor 
               Height          =   360
               Left            =   1032
               TabIndex        =   21
               Top             =   2832
               Width           =   1692
               _Version        =   65537
               _ExtentX        =   2984
               _ExtentY        =   635
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   11
            End
            Begin VB.Label Label7 
               Caption         =   "Conta Corrente"
               Height          =   228
               Left            =   1656
               TabIndex        =   14
               Top             =   1440
               Width           =   1092
            End
            Begin VB.Label Label6 
               Caption         =   "Agência"
               Height          =   228
               Left            =   2112
               TabIndex        =   13
               Top             =   840
               Width           =   612
            End
            Begin VB.Label Label5 
               Caption         =   "Banco"
               Height          =   228
               Left            =   2256
               TabIndex        =   12
               Top             =   240
               Width           =   492
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor"
               Height          =   252
               Left            =   1224
               TabIndex        =   16
               Top             =   2616
               Width           =   1476
            End
            Begin VB.Label Label4 
               Caption         =   "Núm. Cheque"
               Height          =   228
               Left            =   1728
               TabIndex        =   15
               Top             =   2016
               Width           =   996
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Data Nova"
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   744
            TabIndex        =   6
            Top             =   3072
            Width           =   3948
            Begin VB.OptionButton OptdteNova 
               BackColor       =   &H80000000&
               Caption         =   "Nova"
               Height          =   348
               Left            =   300
               TabIndex        =   7
               Top             =   312
               Width           =   1116
            End
            Begin DATEEDITLib.DateEdit DteNova 
               Height          =   348
               Left            =   1968
               TabIndex        =   8
               Top             =   384
               Width           =   1524
               _Version        =   65537
               _ExtentX        =   2688
               _ExtentY        =   614
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Data"
               Height          =   252
               Left            =   2016
               TabIndex        =   9
               Top             =   192
               Width           =   1428
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdInstrucoes 
         Height          =   4956
         Left            =   -74904
         TabIndex        =   4
         Top             =   384
         Width           =   11076
         _ExtentX        =   19553
         _ExtentY        =   8731
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GrdCheque 
         Height          =   4884
         Left            =   -74808
         TabIndex        =   3
         Top             =   528
         Width           =   9324
         _ExtentX        =   16431
         _ExtentY        =   8625
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
      End
   End
End
Attribute VB_Name = "Instrucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            '* Type de Utilização de Banco *'                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Procedures
    Inclusao            As New Custodia.Inserir    'Querys de Insert
    Alteracao           As New Custodia.Atualizar  'Querys de Update
    Deletacao           As New Custodia.Excluir    'Querys de Delete
    Selecao             As New Custodia.Selecionar 'Querys de Select
End Type
Private Procedures      As Procedures

Dim sDtaProcessamento   As Long       'Data de Processamento Formatada MM/DD/AAAA
Dim sDtaAnterior        As Long       'Data Anterior
Dim sDtaNova            As Long       'Data Nova
Dim sValor              As String     'Valor Formatado
Dim sFiltro             As String
Public Sub FormataDataProcessamento()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Processamento *'                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaProcessamento = Mid(Geral.DataProcessamento, 1, 4) & _
                        Mid(Geral.DataProcessamento, 5, 2) & _
                        Mid(Geral.DataProcessamento, 7, 2)

Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data de Processamento.", Err)
    Unload Me

End Sub
Private Sub FormataDataAnterior()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Anterior *'                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaAnterior = Mid(DteAnterior.Text, 5, 4) & _
                   Mid(DteAnterior.Text, 3, 2) & _
                   Mid(DteAnterior.Text, 1, 2)
                   
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data Anterior.", Err)
    Unload Me
                   
End Sub
Private Sub FormataDataNova()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                             '* Formatação da Data Nova *'                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaNova = Mid(DteNova.Text, 5, 4) & _
               Mid(DteNova.Text, 3, 2) & _
               Mid(DteNova.Text, 1, 2)
                   
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data Nova.", Err)
    Unload Me
                   
End Sub
Private Sub TrataPesquisa()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Tratamento de Pesquisa / Formatações  *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsPesquisaInstrucoes    As New ADODB.Recordset
    Dim nContador               As Long

    Screen.MousePointer = vbHourglass

    ''''''''''''''''''''''''''''''''''''
    '   * Verifica e Formata Datas *   '
    ''''''''''''''''''''''''''''''''''''
    If Trim(DteAnterior.Text) <> "" Then
        '''''''''''''''''''''''''''''
        ' * Formata Data Anterior * '
        '''''''''''''''''''''''''''''
        Call FormataDataAnterior
    End If
    
    If Trim(DteNova.Text) <> "" Then
        '''''''''''''''''''''''''''''
        '   * Formata Data Nova *   '
        '''''''''''''''''''''''''''''
        Call FormataDataNova
    End If

    '''''''''''''''''''''''''''''''''''''''''
    '* Verifica e Formata Valor do Cheques *'
    '''''''''''''''''''''''''''''''''''''''''
    If Len(txtValor.Text) = 0 Then
        sValor = 0
    Else
        '''''''''''''''''''''''''''''''''''''''''
        '   *   Formatação de Campo Valor    *  '
        '''''''''''''''''''''''''''''''''''''''''
        sValor = InserePonto(txtValor.Text)
    End If

    ''''''''''''''''''''''''''''''''
    '* Faz Pesquisa de Instruções *'
    ''''''''''''''''''''''''''''''''
    Set rsPesquisaInstrucoes = g_cMainConnection.Execute _
                              (Procedures.Selecao.GetConsultaInstrucaoVC(sDtaProcessamento _
                             , txtnumBordero.Text _
                             , sDtaAnterior _
                             , sDtaNova _
                             , TxtcBco.Text _
                             , txtcAgencia.Text _
                             , TxtCcc.Text _
                             , txtNumCheque.Text _
                             , sValor))
    
    If Not rsPesquisaInstrucoes.EOF Then
    
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdCheque.Rows = rsPesquisaInstrucoes.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        
        GrdInstrucoes.Rows = 1
        
        Do While Not rsPesquisaInstrucoes.EOF
            ''''''''''''''''''''''''''''''''''''''''
            '   * Preenche Grade de Instruções *   '
            ''''''''''''''''''''''''''''''''''''''''
            With GrdInstrucoes
                .Rows = .Rows + 1
                .TextMatrix(nContador, 0) = rsPesquisaInstrucoes!Num_Bordero
                .TextMatrix(nContador, 1) = rsPesquisaInstrucoes!CodigoCarteira
                .TextMatrix(nContador, 2) = Mid(rsPesquisaInstrucoes!DataAnterior, 7, 2) & "/" & Mid(rsPesquisaInstrucoes!DataAnterior, 5, 2) & "/" & Mid(rsPesquisaInstrucoes!DataAnterior, 1, 4)
                .TextMatrix(nContador, 3) = Mid(rsPesquisaInstrucoes!DataNova, 7, 2) & "/" & Mid(rsPesquisaInstrucoes!DataNova, 5, 2) & "/" & Mid(rsPesquisaInstrucoes!DataNova, 1, 4)
                .TextMatrix(nContador, 4) = rsPesquisaInstrucoes!BancoEmitente
                .TextMatrix(nContador, 5) = rsPesquisaInstrucoes!AgenciaEmitente
                .TextMatrix(nContador, 6) = rsPesquisaInstrucoes!CcEmitente
                .TextMatrix(nContador, 7) = rsPesquisaInstrucoes!NrChequeEmitente
                .TextMatrix(nContador, 8) = rsPesquisaInstrucoes!NossoNumero
                .TextMatrix(nContador, 9) = rsPesquisaInstrucoes!CodigoCompensacao
                .TextMatrix(nContador, 10) = Format(rsPesquisaInstrucoes!Valor, "##,##0.00")
            End With
            
            nContador = nContador + 1
            rsPesquisaInstrucoes.MoveNext
            
         Loop
         
        ''''''''''''''''''''''''''''''''''''''''
        '   * Focaliza Tab de Instruções *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(1) = True
        TabOpcoes.Tab = 1
         
        ''''''''''''''''''''''''''''''''''''''''
        '      * Marca 1ª Linha da Grade *     '
        ''''''''''''''''''''''''''''''''''''''''
        Call GrdInstrucoes_SelChange
        Screen.MousePointer = vbDefault
        
    Else
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        Screen.MousePointer = vbDefault
        

        If optbordero.Value Then
               txtnumBordero.SetFocus
               With txtnumBordero
                    .SelStart = 0
                    .SelLength = Len(txtnumBordero.Text)
               End With
               Exit Sub
        Else
          If optdteAnterior.Value Then
               DteAnterior.SetFocus
          Else
               If OptdteNova.Value Then
                    DteNova.SetFocus
               Else
                    If optcheque.Value Then
                         If Len(Trim(TxtcBco.Text)) <> 0 Then
                              TxtcBco.SetFocus
                         ElseIf Len(Trim(txtcAgencia.Text)) <> 0 Then
                              txtcAgencia.SetFocus
                         ElseIf Len(Trim(TxtCcc.Text)) <> 0 Then
                              TxtCcc.SetFocus
                         ElseIf Len(Trim(txtNumCheque.Text)) <> 0 Then
                              txtNumCheque.SetFocus
                         ElseIf Val(txtValor.Text) > 0 Then
                              txtValor.SetFocus
                         End If
                    End If
               End If
          End If
        End If
        
        Exit Sub
    End If
    
    sDtaAnterior = Empty
    sDtaNova = Empty
    
Exit Sub

TrataErro:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao tratar Pesquisas.", Err)
    Unload Me

End Sub
Private Sub cmdConfirmar_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Tratamento de Pesquisa / Formatações  *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''
    '* Pesquisa Por Bco+Ag+Cc+Vlr *'
    ''''''''''''''''''''''''''''''''
    
    If Len(Trim(txtnumBordero.Text)) <> 0 Then
        txtnumBordero = Format(txtnumBordero, String(19, "0"))
    End If

    If optcheque.Value = True Then
        If Len(Trim(TxtcBco.Text)) <> 0 Or Len(Trim(txtcAgencia.Text)) <> 0 Or _
            Len(Trim(TxtCcc.Text)) <> 0 Or Len(Trim(txtNumCheque.Text)) <> 0 Or Len(Trim(txtValor.Text)) <> 0 Then
        Else
            MsgBox "Informe o Banco, Agência, Conta, Nro. Cheque ou Valor a ser pesquisado.", vbInformation + vbOKOnly, App.Title
            TxtcBco.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''
    '* Pesquisa por Numero do Bordero *'
    ''''''''''''''''''''''''''''''''''''
    If optbordero.Value = True Then
        If Len(Trim(txtnumBordero.Text)) = 0 Then
            MsgBox "Informe o número do Borderô a ser pesquisado.", vbInformation + vbOKOnly, App.Title
            txtnumBordero.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''
    '*   Pesquisa por Data Nova   *'
    ''''''''''''''''''''''''''''''''
    If OptdteNova.Value = True Then
        If Len(Trim(DteNova.Text)) = 0 Then
            MsgBox "Informe a Data Nova a ser pesquisada.", vbInformation + vbOKOnly, App.Title
            DteNova.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''
    '* Pesquisa por Data Anterior *'
    ''''''''''''''''''''''''''''''''
    If optdteAnterior.Value = True Then
        If Len(Trim(DteAnterior.Text)) = 0 Then
            MsgBox "Informe a Data Anterior a ser pesquisada.", vbInformation + vbOKOnly, App.Title
            DteAnterior.SetFocus
            Exit Sub
        End If
    End If
    
    Call TrataPesquisa
    
    ''''''''''''''''''''''''''''''''
    '* Valor Default de Variáveis *'
    ''''''''''''''''''''''''''''''''
    sDtaAnterior = 0
    sDtaNova = 0
    
End Sub

Private Sub cmdNovaSelecao_Click()

     '''''''''''''''''''''''''''''''''
     '* Desabilita demais pesquisas *'
     '''''''''''''''''''''''''''''''''
     cmdNovaSelecao.Enabled = False
     optbordero.Value = False
     txtnumBordero.Locked = True
     optcheque.Value = False
     TxtcBco.Locked = True
     txtcAgencia.Locked = True
     TxtCcc.Locked = True
     txtNumCheque.Locked = True
     txtValor.Locked = True
     OptdteNova.Value = False
     DteNova.Locked = True
     optdteAnterior.Value = False
     DteAnterior.Locked = True
    
     '''''''''''''''''''''''''''
     '* Limpa todos os campos *'
     '''''''''''''''''''''''''''
     LimpaTela Me
     
     TabOpcoes.SetFocus

End Sub

Private Sub cmdPrint_Click()

Screen.MousePointer = vbHourglass

If sFiltro <> "" Then
   sFiltro = sFiltro + " and {AlteracaoData.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
Else
   sFiltro = "{AlteracaoData.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
End If

CrystalReport1.ReportFileName = App.path + "\Reports\RelConsultaVC.rpt"
CrystalReport1.WindowTitle = "Consulta de Instruções VC"
CrystalReport1.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
CrystalReport1.Formulas(1) = "CNPJTerceira = '" + Trim(Str(g_Parametros.CNPJ_Terceira)) + "'"
CrystalReport1.SelectionFormula = sFiltro
CrystalReport1.WindowState = crptMaximized
CrystalReport1.WindowTitle = "Relatório de Instruções VC"
CrystalReport1.Action = 0
   
Screen.MousePointer = Default


End Sub

Private Sub CmdSair_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   *  Finaliza Tela  *                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub

Private Sub DteAnterior_GotFocus()
     
     With DteAnterior
          .SelStart = 0
          .SelLength = Len(DteAnterior.Text)
     End With

End Sub

Private Sub DteAnterior_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        cmdConfirmar.SetFocus
        
        ''''''''''''''''''''''''
        '* Se data for válida *'
        ''''''''''''''''''''''''
        If Len(Trim(DteAnterior.Text)) <> 0 Then
            Call cmdConfirmar_Click
        End If
    End If
    
End Sub

Private Sub DteAnterior_LostFocus()

sFiltro = "{AlteracaoData.DataAnterior}=" + Right(DteAnterior.Text, 4) + Mid(DteAnterior.Text, 3, 2) + Left(DteAnterior.Text, 2)

End Sub

Private Sub DteNova_GotFocus()
     
     With DteNova
          .SelStart = 0
          .SelLength = Len(DteNova.Text)
     End With

End Sub

Private Sub DteNova_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        cmdConfirmar.SetFocus
        
        ''''''''''''''''''''''''
        '* Se data for válida *'
        ''''''''''''''''''''''''
        If Len(Trim(DteNova.Text)) <> 0 Then
            Call cmdConfirmar_Click
        End If
    End If
    
End Sub

Private Sub DteNova_LostFocus()

sFiltro = "{AlteracaoData.DataNova}=" + Right(DteNova.Text, 4) + Mid(DteNova.Text, 3, 2) + Left(DteNova.Text, 2)

End Sub

Private Sub Form_Activate()

'Centraliza o form
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

sFiltro = ""

cmdNovaSelecao_Click
     
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                * Carga de Objetos e Formatações *                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    '''''''''''''''''''''''''''''
    ' *  Formatação de Grade  * '
    '''''''''''''''''''''''''''''
    With GrdInstrucoes
        .Cols = 11
        
        ''''''''''''''''''''''''''''''''''
        '   * Nome das Colunas Fixas *   '
        ''''''''''''''''''''''''''''''''''
        .TextMatrix(0, 0) = "Número Borderô"
        .TextMatrix(0, 1) = "Carteira"
        .TextMatrix(0, 2) = "Data Anterior"
        .TextMatrix(0, 3) = "Data Nova"
        .TextMatrix(0, 4) = "Banco"
        .TextMatrix(0, 5) = "Agencia"
        .TextMatrix(0, 6) = "C/C"
        .TextMatrix(0, 7) = "Núm. Cheque"
        .TextMatrix(0, 8) = "Nosso Número"
        .TextMatrix(0, 9) = "Comp."
        .TextMatrix(0, 10) = "Valor"
        
        ''''''''''''''''''''''''''''''''''
        '   * Tamanho de Cada Coluna *   '
        ''''''''''''''''''''''''''''''''''
        .ColWidth(0) = .ColWidth(0) * 1.9
        .ColWidth(1) = .ColWidth(1) * 0.8
        .ColWidth(2) = .ColWidth(2) * 1.1
        .ColWidth(3) = .ColWidth(3) * 1.1
        .ColWidth(4) = .ColWidth(4) * 0.7
        .ColWidth(5) = .ColWidth(5) * 0.8
        .ColWidth(6) = .ColWidth(6) * 1#
        .ColWidth(7) = .ColWidth(7) * 1.2
        .ColWidth(8) = .ColWidth(8) * 1.3
        .ColWidth(9) = .ColWidth(9) * 0.6
        .ColWidth(10) = .ColWidth(10) * 1.3
        
        '''''''''''''''''''''''''''''''''
        '*Alinhamento das Colunas Fixas*'
        '''''''''''''''''''''''''''''''''
        .ColAlignment(0) = 3
        .ColAlignment(1) = 3
        .ColAlignment(2) = 3
        .ColAlignment(3) = 3
        .ColAlignment(4) = 3
        .ColAlignment(5) = 3
        .ColAlignment(6) = 3
        .ColAlignment(7) = 3
        .ColAlignment(8) = 3
        .ColAlignment(9) = 3
        .ColAlignment(10) = 3
    End With

    '''''''''''''''''''''''''''''
    ' *  Formatação das Tabs  * '
    '''''''''''''''''''''''''''''
    TabOpcoes.TabEnabled(1) = False

    '''''''''''''''''''''''''''''''''''
    '* Formata Data de Processamento *'
    '''''''''''''''''''''''''''''''''''
    Call FormataDataProcessamento

Exit Sub
TrataErro:
    Call TratamentoErro("Erro Inicializar Formulário.", Err)
    Unload Me

End Sub

Private Sub GrdInstrucoes_SelChange()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                 * Seleciona Linha Inteira *'                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static m_SelChange As Boolean
    
    'Controle de Acesso
    If m_SelChange = True Then Exit Sub
    
    m_SelChange = True
        
        GrdInstrucoes.Row = GrdInstrucoes.RowSel
        GrdInstrucoes.Col = 0
        GrdInstrucoes.ColSel = 10
        GrdInstrucoes.SetFocus
        
    m_SelChange = False

End Sub
Private Sub optbordero_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       * Habilita Pesquisa por número de Borderô *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     '''''''''''''''''''''''''''''''''
     '* Desabilita demais pesquisas *'
     '''''''''''''''''''''''''''''''''
     cmdNovaSelecao.Enabled = True
     optcheque.Value = False
     TxtcBco.Locked = True
     txtcAgencia.Locked = True
     TxtCcc.Locked = True
     txtNumCheque.Locked = True
     txtValor.Locked = True
     OptdteNova.Value = False
     DteNova.Locked = True
     optdteAnterior.Value = False
     DteAnterior.Locked = True
     
     '''''''''''''''''''''''''''''''''''''''''''''
     '* Habilita Pesquisa por Número de Borderô *'
     '''''''''''''''''''''''''''''''''''''''''''''
     txtnumBordero.Locked = False
         
     '''''''''''''''''''''''''''
     '* Limpa todos os campos *'
     '''''''''''''''''''''''''''
     LimpaTela Me
     
     txtnumBordero.SetFocus
     
End Sub
Private Sub optcheque_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         * Habilita Pesquisa por Dados do Cheque *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     '''''''''''''''''''''''''''''''''
     '* Desabilita demais pesquisas *'
     '''''''''''''''''''''''''''''''''
     cmdNovaSelecao.Enabled = True
     optbordero.Value = False
     txtnumBordero.Locked = True
     OptdteNova.Value = False
     DteNova.Locked = True
     optdteAnterior.Value = False
     DteAnterior.Locked = True
     
     '''''''''''''''''''''''''''''''''''''''''''''
     ' * Habilita Pesquisa por Dados do Cheque * '
     '''''''''''''''''''''''''''''''''''''''''''''
     TxtcBco.Locked = False
     txtcAgencia.Locked = False
     TxtCcc.Locked = False
     txtNumCheque.Locked = False
     txtValor.Locked = False
     
     '''''''''''''''''''''''''''
     '* Limpa todos os campos *'
     '''''''''''''''''''''''''''
     LimpaTela Me
     
     TxtcBco.SetFocus
     
End Sub
Private Sub optdteAnterior_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          * Habilita Pesquisa por Data Anterior *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     '''''''''''''''''''''''''''''''''
     '* Desabilita demais pesquisas *'
     '''''''''''''''''''''''''''''''''
     cmdNovaSelecao.Enabled = True
     optcheque.Value = False
     TxtcBco.Locked = True
     txtcAgencia.Locked = True
     TxtCcc.Locked = True
     txtNumCheque.Locked = True
     txtValor.Locked = True
     optbordero.Value = False
     txtnumBordero.Locked = True
     OptdteNova.Value = False
     DteNova.Locked = True
     
     '''''''''''''''''''''''''''''''''''''''''''
     ' * Habilita Pesquisa por Data Anterior * '
     '''''''''''''''''''''''''''''''''''''''''''
     DteAnterior.Locked = False
     
     '''''''''''''''''''''''''''
     '* Limpa todos os campos *'
     '''''''''''''''''''''''''''
     LimpaTela Me
     
     DteAnterior.SetFocus
     
End Sub
Private Sub OptdteNova_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            * Habilita Pesquisa por Data Nova *                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     '''''''''''''''''''''''''''''''''
     '* Desabilita demais pesquisas *'
     '''''''''''''''''''''''''''''''''
     cmdNovaSelecao.Enabled = True
     optcheque.Value = False
     TxtcBco.Locked = True
     txtcAgencia.Locked = True
     TxtCcc.Locked = True
     txtNumCheque.Locked = True
     txtValor.Locked = True
     optbordero.Value = False
     txtnumBordero.Locked = True
     optdteAnterior.Value = False
     DteAnterior.Locked = True
    
     '''''''''''''''''''''''''''''''''''''''''''
     ' * Habilita Pesquisa por Data Anterior * '
     '''''''''''''''''''''''''''''''''''''''''''
     DteNova.Locked = False
    
     '''''''''''''''''''''''''''
     '* Limpa todos os campos *'
     '''''''''''''''''''''''''''
     LimpaTela Me
    
     DteNova.SetFocus
     
End Sub
Private Sub TabOpcoes_Click(PreviousTab As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       *  Se clicar na Tab 0 desabilita as demais *                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     If TabOpcoes.Tab = 0 Then
        
          TabOpcoes.TabEnabled(0) = True
          TabOpcoes.TabEnabled(1) = False
               
          ''''''''''''''''''''''''
          ' * Pesquisa Default * '
          ''''''''''''''''''''''''
          If optbordero.Value Then
               Call optbordero_Click
               optbordero.Value = True
          ElseIf optcheque.Value Then
               Call optcheque_Click
          ElseIf optdteAnterior.Value Then
               Call optdteAnterior_Click
          ElseIf OptdteNova.Value Then
               Call OptdteNova_Click
          End If
          
          cmdPrint.Visible = False
          cmdConfirmar.Visible = True
          cmdNovaSelecao.Visible = True
     
     Else
          cmdPrint.Visible = True
          cmdConfirmar.Visible = False
          cmdNovaSelecao.Visible = False
     End If
        
End Sub

Private Sub txtcAgencia_GotFocus()

     With txtcAgencia
          .SelStart = 0
          .SelLength = Len(txtcAgencia)
     End With

End Sub

Private Sub txtcAgencia_KeyPress(KeyAscii As Integer)
    
    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtcAgencia_LostFocus()

If sFiltro = "" Then
   sFiltro = "{AlteracaoData.AgenciaEmitente}=" + Trim(txtcAgencia.Text)
Else
   sFiltro = sFiltro + " and {AlteracaoData.AgenciaEmitente}=" + Trim(txtcAgencia.Text)
End If

End Sub

Private Sub TxtcBco_GotFocus()

     With TxtcBco
          .SelStart = 0
          .SelLength = Len(TxtcBco)
     End With

End Sub

Private Sub TxtcBco_KeyPress(KeyAscii As Integer)
    
    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtcBco_LostFocus()

sFiltro = "{AlteracaoData.BancoEmitente}=" + Trim(TxtcBco.Text)

End Sub

Private Sub TxtCcc_GotFocus()
     
     With TxtCcc
          .SelStart = 0
          .SelLength = Len(TxtCcc)
     End With

End Sub

Private Sub TxtCcc_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtCcc_LostFocus()

If sFiltro = "" Then
   sFiltro = "{AlteracaoData.CCEmitente}=" + Trim(TxtCcc.Text)
Else
   sFiltro = sFiltro + " and {AlteracaoData.CCEmitente}=" + Trim(TxtCcc.Text)
End If

End Sub

Private Sub txtnumBordero_GotFocus()

     With txtnumBordero
          .SelStart = 0
          .SelLength = Len(txtnumBordero.Text)
     End With
     
End Sub

Private Sub txtnumBordero_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        '''''''''''''''''''''''''''''''''''''''''''''''
        '* Formata número do borderô com 19 posições *'
        '''''''''''''''''''''''''''''''''''''''''''''''
        Call cmdConfirmar_Click
    End If
    
End Sub

Private Sub txtnumBordero_LostFocus()

sFiltro = "{AlteracaoData.Num_Bordero}='" + Trim(txtnumBordero.Text) + "'"

End Sub

Private Sub txtNumCheque_GotFocus()

     With txtNumCheque
          .SelStart = 0
          .SelLength = Len(txtNumCheque)
     End With

End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtNumCheque_LostFocus()

If sFiltro = "" Then
   sFiltro = "{AlteracaoData.NrChequeEmitente}=" + Trim(TxtCcc.Text)
Else
   sFiltro = sFiltro + " and {AlteracaoData.NrChequeEmitente}=" + Trim(TxtCcc.Text)
End If

End Sub

Private Sub txtValor_GotFocus()
     
     With txtValor
          .SelStart = 0
          .SelLength = Len(txtValor.Text)
     End With

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

    ''''''''''''''''''''''''''''''''''
    '* Este campo só aceita números *'
    ''''''''''''''''''''''''''''''''''
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        Call cmdConfirmar_Click
    End If
    
End Sub

Private Sub txtValor_LostFocus()

If sFiltro = "" Then
   sFiltro = "{AlteracaoData.Valor}=" + Trim(txtValor.Text)
Else
   sFiltro = sFiltro + " and {AlteracaoData.Valor}=" + Trim(txtValor.Text)
End If

End Sub
