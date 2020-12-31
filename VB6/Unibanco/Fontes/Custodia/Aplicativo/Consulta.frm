VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   6660
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   720
      Left            =   7200
      Picture         =   "Consulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5856
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   720
      Left            =   8064
      Picture         =   "Consulta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5856
      Width           =   816
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3480
      Top             =   6000
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab TabOpcoes 
      Height          =   5652
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   72
      Width           =   9684
      _ExtentX        =   17082
      _ExtentY        =   9970
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "&Argumentos de Pesquisa"
      TabPicture(0)   =   "Consulta.frx":0614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frafavorecido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fracheque"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDeposito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frabordero"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frastatus"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Cheques"
      TabPicture(1)   =   "Consulta.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdCheque"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Borderôs"
      TabPicture(2)   =   "Consulta.frx":064C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtStatus"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtNmCliente"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "GrdDatasDeposito"
      Tab(2).Control(3)=   "GrdBordero"
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(5)=   "Label4"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   -70080
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   744
         Width           =   4572
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Valor do Cheque"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   696
         TabIndex        =   35
         Top             =   3552
         Width           =   3948
         Begin CURRENCYEDITLib.CurrencyEdit txtValor 
            Height          =   348
            Left            =   1920
            TabIndex        =   10
            Top             =   408
            Width           =   1836
            _Version        =   65537
            _ExtentX        =   3238
            _ExtentY        =   614
            _StockProps     =   93
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
         End
         Begin VB.OptionButton OptValor 
            BackColor       =   &H80000000&
            Caption         =   "Valor"
            Height          =   348
            Left            =   336
            TabIndex        =   9
            Top             =   312
            Width           =   1116
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Valor"
            Height          =   252
            Left            =   1968
            TabIndex        =   36
            Top             =   192
            Width           =   1716
         End
      End
      Begin VB.TextBox txtNmCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   -74736
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   744
         Width           =   4572
      End
      Begin MSFlexGridLib.MSFlexGrid GrdCheque 
         Height          =   4884
         Left            =   -74832
         TabIndex        =   16
         Top             =   528
         Width           =   9324
         _ExtentX        =   16447
         _ExtentY        =   8615
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
      End
      Begin VB.Frame frastatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Status do Borderô"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   5016
         TabIndex        =   25
         Top             =   2352
         Width           =   3948
         Begin VB.ComboBox CboStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   384
            Width           =   2484
         End
         Begin VB.OptionButton optstatus 
            BackColor       =   &H80000000&
            Caption         =   "Status"
            Height          =   348
            Left            =   270
            TabIndex        =   7
            Top             =   360
            Width           =   1116
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Status"
            Height          =   252
            Left            =   1344
            TabIndex        =   26
            Top             =   192
            Width           =   2412
         End
      End
      Begin VB.Frame frabordero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Identificação do Borderô"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   5016
         TabIndex        =   24
         Top             =   1152
         Width           =   3948
         Begin VB.TextBox txtnumBordero 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            MaxLength       =   19
            TabIndex        =   4
            Top             =   384
            Width           =   2484
         End
         Begin VB.OptionButton optbordero 
            BackColor       =   &H80000000&
            Caption         =   "Borderô"
            Height          =   348
            Left            =   270
            TabIndex        =   3
            Top             =   312
            Width           =   972
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Identificação"
            Height          =   252
            Left            =   1488
            TabIndex        =   27
            Top             =   192
            Width           =   2172
         End
      End
      Begin VB.Frame fraDeposito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Datas de Depósito"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   696
         TabIndex        =   22
         Top             =   2352
         Width           =   3948
         Begin DATEEDITLib.DateEdit DteDtaDeposito 
            Height          =   348
            Left            =   1968
            TabIndex        =   6
            Top             =   384
            Width           =   1524
            _Version        =   65537
            _ExtentX        =   2688
            _ExtentY        =   614
            _StockProps     =   93
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
         Begin VB.OptionButton optdeposito 
            BackColor       =   &H80000000&
            Caption         =   "Depósito"
            Height          =   348
            Left            =   312
            TabIndex        =   5
            Top             =   312
            Width           =   1116
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Data"
            Height          =   252
            Left            =   2016
            TabIndex        =   23
            Top             =   192
            Width           =   1260
         End
      End
      Begin VB.Frame fracheque 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Identificação do Cheque"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   5040
         TabIndex        =   21
         Top             =   3552
         Width           =   3948
         Begin VB.TextBox TxtCcc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2664
            MaxLength       =   10
            TabIndex        =   14
            Top             =   420
            Width           =   1092
         End
         Begin VB.TextBox txtcAgencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1950
            MaxLength       =   4
            TabIndex        =   13
            Top             =   420
            Width           =   516
         End
         Begin VB.TextBox txtcBco 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1350
            MaxLength       =   3
            TabIndex        =   12
            Top             =   420
            Width           =   444
         End
         Begin VB.OptionButton optcheque 
            BackColor       =   &H80000000&
            Caption         =   "Cheque"
            Height          =   348
            Left            =   270
            TabIndex        =   11
            Top             =   312
            Width           =   945
         End
         Begin VB.Label Label7 
            Caption         =   "Conta Corrente"
            Height          =   228
            Left            =   2688
            TabIndex        =   32
            Top             =   192
            Width           =   1092
         End
         Begin VB.Label Label6 
            Caption         =   "Agência"
            Height          =   228
            Left            =   1920
            TabIndex        =   31
            Top             =   192
            Width           =   636
         End
         Begin VB.Label Label5 
            Caption         =   "Banco"
            Height          =   228
            Left            =   1296
            TabIndex        =   30
            Top             =   192
            Width           =   492
         End
      End
      Begin VB.Frame frafavorecido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Dados do Favorecido"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   696
         TabIndex        =   20
         Top             =   1152
         Width           =   3948
         Begin VB.TextBox txtfagencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1968
            MaxLength       =   4
            TabIndex        =   1
            Top             =   384
            Width           =   516
         End
         Begin VB.TextBox Txtfcc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   2
            Top             =   384
            Width           =   1092
         End
         Begin VB.OptionButton optfavorecido 
            BackColor       =   &H80000000&
            Caption         =   "Favorecido"
            Height          =   348
            Left            =   312
            TabIndex        =   0
            Top             =   312
            Width           =   1116
         End
         Begin VB.Label Label9 
            Caption         =   "Conta Corrente"
            Height          =   228
            Left            =   2664
            TabIndex        =   34
            Top             =   192
            Width           =   1092
         End
         Begin VB.Label Label8 
            Caption         =   "Agência"
            Height          =   228
            Left            =   1944
            TabIndex        =   33
            Top             =   192
            Width           =   636
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDatasDeposito 
         Height          =   1908
         Left            =   -74808
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3528
         Width           =   9324
         _ExtentX        =   16447
         _ExtentY        =   3366
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid GrdBordero 
         Height          =   2148
         Left            =   -74808
         TabIndex        =   17
         Top             =   1200
         Width           =   9324
         _ExtentX        =   16447
         _ExtentY        =   3789
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Status do Borderô :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   -70008
         TabIndex        =   38
         Top             =   504
         Width           =   4428
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Nome do Cliente :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   -74664
         TabIndex        =   29
         Top             =   504
         Width           =   4428
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   720
      Left            =   8976
      Picture         =   "Consulta.frx":0668
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5856
      Width           =   816
   End
End
Attribute VB_Name = "Consulta"
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

Dim sDtaProcessamento   As Long
Dim sDtaDeposito        As Long

Dim sDataDeposito       As String
Dim sStatusDefault      As String
Dim sFiltroBordero      As String
Dim sFiltroCheque       As String
Private Function FormataCpfCnpj(ByVal CNPJCPF As String) As String

Dim sCodigo As String

sCodigo = Trim(CNPJCPF)

'Verifica se é CPF ou CNPJ
If Len(sCodigo) > 11 Then
     sCodigo = Right(String(14, "0") & Trim(CNPJCPF), 14)
     FormataCpfCnpj = Format(sCodigo, "@@.@@@.@@@/@@@@-@@")
Else
     sCodigo = Right(String(11, "0") & Trim(CNPJCPF), 11)
     FormataCpfCnpj = Format(sCodigo, "@@@.@@@.@@@-@@")
End If

End Function
Private Sub ListaStatusBordero()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Traz a Lista de Status do Bordero *'                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsListaCheque As New ADODB.Recordset

    Set rsListaCheque = g_cMainConnection.Execute(Procedures.Selecao.GetStatusBordero)

    If Not rsListaCheque.EOF Then
        
        '''''''''''''''''''''''''''''''''''''''''
        '     * Default do Combo de Status *    '
        '''''''''''''''''''''''''''''''''''''''''
        sStatusDefault = rsListaCheque!Descricao
        'CboStatus.Text = rsListaCheque!Descricao
        
        Do While Not rsListaCheque.EOF
            '''''''''''''''''''''''''''
            ' * Descrição do Status * '
            '''''''''''''''''''''''''''
            CboStatus.AddItem rsListaCheque!Descricao
            rsListaCheque.MoveNext
        Loop
        
    End If
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao Listar Status do Cheque.", Err)
    Unload Me
End Sub
Private Function PesquisaBordero() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    * Pesquisa de Borderôs *                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsBordero   As New ADODB.Recordset
    Dim nContador   As Integer
    Dim nAgencia    As Integer
    Dim nConta      As Long
    
    Dim Controle    As Control

    PesquisaBordero = False
    
    '''''''''''''''''''''''''''''''
    ' * Define Valor de Agência * '
    '''''''''''''''''''''''''''''''
    If Len(Trim(txtfagencia)) = 0 Then
        nAgencia = 0
    Else
        nAgencia = CInt(txtfagencia)
    End If
 
    '''''''''''''''''''''''''''''''
    '  * Define Valor de Conta *  '
    '''''''''''''''''''''''''''''''
    If Len(Trim(Txtfcc)) = 0 Then
        nConta = 0
    Else
        nConta = CLng(Txtfcc)
    End If
 
    Set rsBordero = g_cMainConnection.Execute _
                   (Procedures.Selecao.GetConsultaBordero _
                   (sDtaProcessamento, _
                   nAgencia, _
                   nConta, _
                   0, _
                   IIf(optstatus.Value = True, CboStatus.Text, "")))

    If Not rsBordero.EOF Then

        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdBordero.Rows = rsBordero.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        If rsBordero.RecordCount > 8 Then
            GrdBordero.ColWidth(5) = GrdBordero.Width * 0.14   'Ajustar largura também no (Form Load) e (PesquisaIdBordero)
        Else
            GrdBordero.ColWidth(5) = GrdBordero.Width * 0.161  'Ajustar largura também no (Form Load) e (PesquisaIdBordero)
        End If
        
        Do While Not rsBordero.EOF
            
            ''''''''''''''''''''''''''''''''''''''
            '   * Preenche Grade de Borderôs *   '
            ''''''''''''''''''''''''''''''''''''''
            With GrdBordero
                .TextMatrix(nContador, 0) = Format(rsBordero!Num_Bordero, String(19, "0"))
                .TextMatrix(nContador, 1) = Format(rsBordero!Agencia, String(4, "0"))
                .TextMatrix(nContador, 2) = Format(rsBordero!Conta, String(7, "0"))
                .TextMatrix(nContador, 3) = rsBordero!Descricao
                .TextMatrix(nContador, 4) = Format(rsBordero!CodigoLoja, String(10, "0"))
                .TextMatrix(nContador, 5) = Mid(rsBordero!DataEntrada, 7, 2) & "/" & Mid(rsBordero!DataEntrada, 5, 2) & "/" & Mid(rsBordero!DataEntrada, 1, 4)
                .TextMatrix(nContador, 6) = rsBordero!NomeCliente
                .TextMatrix(nContador, 7) = rsBordero!IdBordero
                .TextMatrix(nContador, 8) = rsBordero!Status
            End With
            
            nContador = nContador + 1
            rsBordero.MoveNext
            
         Loop
         
         PesquisaBordero = True
         
        ''''''''''''''''''''''''''''''''''''''''
        '     * Focaliza Tab de Borderos *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(2) = True
        TabOpcoes.Tab = 2

    Else
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        
         For Each Controle In Consulta.Controls
        
            If TypeName(Controle) = "TextBox" Or TypeName(Controle) = "DateEdit" Then
               
               If Len(Controle.Text) <> 0 Then
                  Controle.SelStart = 0
                  Controle.SelLength = Len(Controle.Text)
                  Controle.SetFocus
                  Exit For
               End If
               
            End If
            
         Next
        
        Exit Function
        
    End If

Exit Function
TrataErro:
    Call TratamentoErro("Erro ao pesquisar Borderô (Favorecido).", Err)
    Unload Me

End Function
Private Sub PesquisaIdBordero()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    * Pesquisa de Borderôs *                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsBordero   As New ADODB.Recordset
    Dim nContador   As Integer
    
    Dim Controle    As Control

    Set rsBordero = g_cMainConnection.Execute(Procedures.Selecao.GetConsultaBordero _
                                             (sDtaProcessamento _
                                             , _
                                             , _
                                             , GrdCheque.TextMatrix(GrdCheque.Row, 8)))

    If Not rsBordero.EOF Then

        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdBordero.Rows = rsBordero.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        If rsBordero.RecordCount > 8 Then
             GrdBordero.ColWidth(5) = GrdBordero.Width * 0.14   'Ajustar largura também no (Form Load) e (PesquisaBordero)
        Else
             GrdBordero.ColWidth(5) = GrdBordero.Width * 0.161   'Ajustar largura também no (Form Load) e (PesquisaBordero)
        End If
        
        Do While Not rsBordero.EOF
            
            ''''''''''''''''''''''''''''''''''''''
            '   * Preenche Grade de Borderôs *   '
            ''''''''''''''''''''''''''''''''''''''
            With GrdBordero
                .TextMatrix(nContador, 0) = Format(rsBordero!Num_Bordero, String(19, "0"))
                .TextMatrix(nContador, 1) = Format(rsBordero!Agencia, String(4, "0"))
                .TextMatrix(nContador, 2) = Format(rsBordero!Conta, String(7, "0"))
                .TextMatrix(nContador, 3) = rsBordero!Descricao
                .TextMatrix(nContador, 4) = Format(rsBordero!CodigoLoja, String(10, "0"))
                .TextMatrix(nContador, 5) = Mid(rsBordero!DataEntrada, 7, 2) & "/" & Mid(rsBordero!DataEntrada, 5, 2) & "/" & Mid(rsBordero!DataEntrada, 1, 4)
                .TextMatrix(nContador, 6) = rsBordero!NomeCliente
                .TextMatrix(nContador, 7) = rsBordero!IdBordero
                .TextMatrix(nContador, 8) = rsBordero!Status
            End With
            
            nContador = nContador + 1
            rsBordero.MoveNext
            
         Loop
         
        ''''''''''''''''''''''''''''''''''''''''
        '     * Focaliza Tab de Borderos *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(2) = True
        TabOpcoes.Tab = 2
         
        ''''''''''''''''''''''''''''''''''''''''
        '     * Marca a 1ª Linha da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        Call GrdCheque_SelChange
    Else
    
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        
         For Each Controle In Me.Controls
        
            If TypeName(Controle) = "TextBox" Or TypeName(Controle) = "DateEdit" Then
            
               If Len(Controle.Text) <> 0 Then
                  Controle.SelStart = 0
                  Controle.SelLength = Len(Controle.Text)
                  Controle.SetFocus
                  Exit For
               End If
               
           End If
            
        Next
        
        Exit Sub
        
    End If
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao pesquisar Borderô.", Err)
    Unload Me
    
End Sub
Private Function PesquisaCheque() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    * Pesquisa de Cheques *                                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsCheques   As New ADODB.Recordset
    Dim nContador   As Integer
    Dim Valor       As String
    
    Dim Controle    As Control

    PesquisaCheque = False
    
    '''''''''''''''''''''''''''''''''''''''''
    '* Verifica e Formata Valor do Cheques *'
    '''''''''''''''''''''''''''''''''''''''''
    If Len(txtValor.Text) = 0 Then
        Valor = 0
    Else
        '''''''''''''''''''''''''''''''''''''''''
        '   *   Formatação de Campo Valor    *  '
        '''''''''''''''''''''''''''''''''''''''''
        Valor = InserePonto(txtValor.Text)
    End If

    Set rsCheques = g_cMainConnection.Execute(Procedures.Selecao.GetConsultaCheque _
                                             (sDtaProcessamento _
                                            , sDtaDeposito _
                                            , IIf(optstatus.Value = True, CboStatus.Text, "") _
                                            , txtcBco.Text _
                                            , txtcAgencia.Text _
                                            , TxtCcc.Text _
                                            , Valor _
                                            ))
            
    If Not rsCheques.EOF Then

        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdCheque.Rows = rsCheques.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        If rsCheques.RecordCount > 20 Then
             GrdCheque.ColWidth(7) = GrdCheque.Width * 0.142   'Ajustar largura também no (Form Load) e (PesquisaChequesBordero)
             GrdCheque.ColWidth(6) = GrdCheque.Width * 0.142
        Else
             GrdCheque.ColWidth(7) = GrdCheque.Width * 0.162   'Ajustar largura também no (Form Load) e (PesquisaChequesBordero)
             GrdCheque.ColWidth(6) = GrdCheque.Width * 0.162
        End If
        
        Do While Not rsCheques.EOF
            ''''''''''''''''''''''''''''''''''''''''
            '     * Preenche Grade de Cheques *    '
            ''''''''''''''''''''''''''''''''''''''''
            GrdCheque.ForeColor = vbBlack
            With GrdCheque
                .TextMatrix(nContador, 0) = rsCheques!Banco
                .TextMatrix(nContador, 1) = rsCheques!Agencia
                .TextMatrix(nContador, 2) = rsCheques!ContaC
                .TextMatrix(nContador, 3) = rsCheques!NrCheque
                .TextMatrix(nContador, 4) = Mid(rsCheques!DataDeposito, 7, 2) & "/" & Mid(rsCheques!DataDeposito, 5, 2) & "/" & Mid(rsCheques!DataDeposito, 1, 4)
                .TextMatrix(nContador, 5) = IIf(IsNull(rsCheques!CNPJCPF), "", FormataCpfCnpj(rsCheques!CNPJCPF))
                .TextMatrix(nContador, 6) = rsCheques!Descricao
                .Row = nContador
                .Col = 6
                If rsCheques!Status = "D" Then
                    .CellForeColor = vbRed
                ElseIf rsCheques!Status = "I" Then
                    .CellForeColor = vbBlue
                Else
                    .CellForeColor = vbBlack
                End If
                
                .TextMatrix(nContador, 7) = Format(rsCheques!Valor, "##,#00.00")
                .TextMatrix(nContador, 8) = rsCheques!IdBordero
            End With
            
            nContador = nContador + 1
            rsCheques.MoveNext
            
         Loop
         
        ''''''''''''''''''''''''''''''''''''''''
        '      * Focaliza Tab de Cheques *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(1) = True
        TabOpcoes.Tab = 1
         
        PesquisaCheque = True
        
        ''''''''''''''''''''''''''''''''''''''''
        '      * Marca 1ª Linha da Grade *     '
        ''''''''''''''''''''''''''''''''''''''''
        Call GrdCheque_SelChange
    Else
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        
         For Each Controle In Me.Controls
        
            If TypeName(Controle) = "TextBox" Or TypeName(Controle) = "DateEdit" Then
               
               If Len(Controle.Text) <> 0 Then
                  Controle.SelStart = 0
                  Controle.SelLength = Len(Controle.Text)
                  Controle.SetFocus
                  Exit For
               End If
               
            End If
            
        Next
        
        Exit Function
        
    End If
    
Exit Function
TrataErro:
    Call TratamentoErro("Erro ao pesquisar Cheques.", Err)
    Unload Me
    
End Function
Private Function PesquisaChequesBordero() As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            * Pesquisa de Cheques do Bordero *                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
    
    Dim rsCheques    As New ADODB.Recordset
    Dim nContador    As Integer
    
    Dim Controle     As Control
    
    PesquisaChequesBordero = False
    
    Set rsCheques = g_cMainConnection.Execute(Procedures.Selecao.GetConsultaChequesBordero _
                    (sDtaProcessamento, (txtnumBordero.Text)))
                    
                                              
    If Not rsCheques.EOF Then

        sFiltroCheque = "{Cheque.IdBordero}=" + Trim(Str(rsCheques!IdBordero))

        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdCheque.Rows = rsCheques.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        If rsCheques.RecordCount > 20 Then
             GrdCheque.ColWidth(7) = GrdCheque.Width * 0.142   'Ajustar largura também no (Form Load) e (PesquisaCheque)
        Else
             GrdCheque.ColWidth(7) = GrdCheque.Width * 0.162   'Ajustar largura também no (Form Load) e (PesquisaCheque)
        End If
        
        Do While Not rsCheques.EOF
            ''''''''''''''''''''''''''''''''''''''''
            '     * Preenche Grade de Cheques *    '
            ''''''''''''''''''''''''''''''''''''''''
            With GrdCheque
                .TextMatrix(nContador, 0) = rsCheques!Banco
                .TextMatrix(nContador, 1) = rsCheques!Agencia
                .TextMatrix(nContador, 2) = rsCheques!ContaC
                .TextMatrix(nContador, 3) = rsCheques!NrCheque
                .TextMatrix(nContador, 4) = Mid(rsCheques!DataDeposito, 7, 2) & "/" & Mid(rsCheques!DataDeposito, 5, 2) & "/" & Mid(rsCheques!DataDeposito, 1, 4)
                .TextMatrix(nContador, 5) = IIf(IsNull(rsCheques!CNPJCPF), "", FormataCpfCnpj(rsCheques!CNPJCPF))
                .TextMatrix(nContador, 6) = rsCheques!Descricao
                .TextMatrix(nContador, 7) = Format(rsCheques!Valor, "##,#00.00")
                .TextMatrix(nContador, 8) = rsCheques!IdBordero
                
                .Row = nContador: .Col = 6
                If rsCheques!Status = "D" Then
                    .CellForeColor = vbRed
                ElseIf rsCheques!Status = "I" Then
                    .CellForeColor = vbBlue
                Else
                    .CellForeColor = vbBlack
                End If
                
            End With
            
            nContador = nContador + 1
            rsCheques.MoveNext
            
         Loop
         
        ''''''''''''''''''''''''''''''''''''''''
        '      * Focaliza Tab de Cheques *     '
        ''''''''''''''''''''''''''''''''''''''''
        TabOpcoes.TabEnabled(1) = True
        TabOpcoes.Tab = 1
         
        PesquisaChequesBordero = True
        
        ''''''''''''''''''''''''''''''''''''''''
        '      * Marca 1ª Linha da Grade *     '
        ''''''''''''''''''''''''''''''''''''''''
        Call GrdCheque_SelChange
    Else
        MsgBox "Nenhum registro foi encontrado.", vbExclamation + vbOKOnly, App.Title
        
         For Each Controle In Me.Controls
        
            If TypeName(Controle) = "TextBox" Or TypeName(Controle) = "DateEdit" Then
               
               If Len(Controle.Text) <> 0 Then
                  Controle.SelStart = 0
                  Controle.SelLength = Len(Controle.Text)
                  Controle.SetFocus
                  Exit For
               End If
               
            End If
            
        Next
        
        Exit Function
        
    End If

Exit Function
TrataErro:
    Call TratamentoErro("Erro ao pesquisar Cheques do Borderô.", Err)
    Unload Me

End Function
Private Sub MontaGradeDeDatas()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'              * Monta a Grade de Datas de Depósito com Base no Número *                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsDatasDeposito As New ADODB.Recordset
    Dim nContador       As Integer

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Mostra o Nome do Cliente do TextBox de acordo com o Borderô * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtNmCliente.Text = UCase(GrdBordero.TextMatrix(GrdBordero.Row, 6))
    
    '''''''''''''''''''''''''''''''''
    ' * Mostra o Status do Borderô *'
    '''''''''''''''''''''''''''''''''
    txtStatus.Text = UCase(GrdBordero.TextMatrix(GrdBordero.Row, 8))
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Pesquisa de Datas de Depósito para o Borderô Atual * '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rsDatasDeposito = g_cMainConnection.Execute(Procedures.Selecao.GetDatasBordero _
                                                   (sDtaProcessamento _
                                                  , GrdBordero.TextMatrix(GrdBordero.Row, 7)))
    
    If Not rsDatasDeposito.EOF Then
    
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número de Linhas da Grade *    '
        ''''''''''''''''''''''''''''''''''''''''
        GrdDatasDeposito.Rows = rsDatasDeposito.RecordCount + 1
        
        ''''''''''''''''''''''''''''''''''''''''
        '     * Número incial do Contador *    '
        ''''''''''''''''''''''''''''''''''''''''
        nContador = 1
        If Not rsDatasDeposito.EOF Then
          If rsDatasDeposito.RecordCount > 7 Then
               GrdDatasDeposito.ColWidth(2) = GrdDatasDeposito.Width * 0.312  'Ajustar largura também no (Form Load)
          Else
               GrdDatasDeposito.ColWidth(2) = GrdDatasDeposito.Width * 0.332  'Ajustar largura também no (Form Load)
          End If
        End If
        
        
        Do While Not rsDatasDeposito.EOF
            ''''''''''''''''''''''''''''''''''''''''
            '     * Preenche Grade de Cheques *    '
            ''''''''''''''''''''''''''''''''''''''''
            With GrdDatasDeposito
                .TextMatrix(nContador, 0) = Mid(rsDatasDeposito!DataDeposito, 7, 2) & "/" & Mid(rsDatasDeposito!DataDeposito, 5, 2) & "/" & Mid(rsDatasDeposito!DataDeposito, 1, 4)
                .TextMatrix(nContador, 1) = rsDatasDeposito!QuantidadeCheques
                .TextMatrix(nContador, 2) = Format(rsDatasDeposito!ValorDeposito, "##,##0.00")
            End With
            
            nContador = nContador + 1
            rsDatasDeposito.MoveNext
            
         Loop
    End If
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao montar Grade de Datas.", Err)
    Unload Me
    
End Sub
Private Sub FormataDataProcessamento()
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
Private Sub FormataDataDeposito()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Depósito *'                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    sDtaDeposito = Mid(DteDtaDeposito.Text, 5, 4) & _
                   Mid(DteDtaDeposito.Text, 3, 2) & _
                   Mid(DteDtaDeposito.Text, 1, 2)
                   
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao formatar Data de Depósito.", Err)
    Unload Me
                   
End Sub

Private Sub CboStatus_GotFocus()

sFiltroBordero = ""

End Sub

Private Sub CboStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboStatus_LostFocus
        cmdConfirmar_Click
        Exit Sub
    End If
End Sub

Private Sub CboStatus_LostFocus()

sFiltroBordero = "{StatusBordero.Descricao}='" + Trim(CboStatus.Text) + "'"

End Sub

Private Sub cmdConfirmar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   * Trata Tipos de Pesquisa *                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
    '''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa de Cheques do com Nº Borderô *'
    '''''''''''''''''''''''''''''''''''''''''''
    If optbordero.Value = True Then
        If Len(Trim(txtnumBordero.Text)) <> 0 Then
            txtnumBordero.Text = Format(CStr(txtnumBordero.Text), String(19, "0"))
            If Not PesquisaChequesBordero Then Exit Sub
            
        Else
            MsgBox "Informe o número do Borderô.", vbInformation + vbOKOnly, App.Title
            txtnumBordero.SetFocus
            Exit Sub
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa de Cheques do com: Bco+Ag+Cc *'
    '''''''''''''''''''''''''''''''''''''''''''
    If optcheque.Value = True Then
        If Len(Trim(txtcBco.Text)) <> 0 Or Len(Trim(txtcAgencia.Text)) <> 0 Or Len(Trim(TxtCcc.Text)) <> 0 Then
            If Not PesquisaCheque Then Exit Sub
        Else
            MsgBox "Informe Banco, Agência ou Conta do Cheque.", vbInformation + vbOKOnly, App.Title
            txtcBco.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa Datas de Depósito do Cheque *'
    ''''''''''''''''''''''''''''''''''''''''''
    If optdeposito.Value = True Then
        If Len(Trim(DteDtaDeposito.Text)) <> 0 Then
            Call FormataDataDeposito
            If Not PesquisaCheque Then Exit Sub
            ''''''''''''''''''''''''''''''''''
            '     * Limpeza de Variavel *    '
            ''''''''''''''''''''''''''''''''''
            sDtaDeposito = 0
        Else
            MsgBox "Informe a Data de Depósito.", vbInformation + vbOKOnly, App.Title
            DteDtaDeposito.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa Borderô do Favorecio: Ag+Cc *'
    ''''''''''''''''''''''''''''''''''''''''''
    If optfavorecido.Value = True Then
        If Len(Trim(txtfagencia.Text)) <> 0 Or Len(Trim(Txtfcc.Text)) <> 0 Then
            If Not PesquisaBordero Then Exit Sub
        Else
            MsgBox "Informe Agência ou Conta do Favorecido.", vbInformation + vbOKOnly, App.Title
            txtfagencia.SetFocus
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    '*   Pesquisa Cheques com seu Status    *'
    ''''''''''''''''''''''''''''''''''''''''''
    If optstatus.Value = True Then
        If Not PesquisaBordero Then Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''
    '*     Pesquisa de Cheques por Valor     *'
    '''''''''''''''''''''''''''''''''''''''''''
    If OptValor.Value = True Then
        If Len(Trim(txtValor.Text)) <> 0 Then
            If txtValor.Text <= "99999999999" Then
                If Not PesquisaCheque Then Exit Sub
            Else
                MsgBox "Valor é maior que o permitido.", vbInformation + vbOKOnly, App.Title
                txtValor.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Valor não pode estar zerado.", vbInformation + vbOKOnly, App.Title
            txtValor.SetFocus
            Exit Sub
        End If
    End If
    
    cmdPrint.Enabled = True
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao tratar tipos de pesquisa.", Err)
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()

Screen.MousePointer = vbHourglass

If TabOpcoes.Tab = 1 Then

   If sFiltroCheque <> "" Then
      sFiltroCheque = sFiltroCheque + " and {Cheque.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
   Else
      sFiltroCheque = "{Cheque.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
   End If
   
   CrystalReport1.ReportFileName = App.path + "\Reports\RelConsultaCheques.rpt"
   CrystalReport1.WindowTitle = "Consulta de Cheques"
   CrystalReport1.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
   CrystalReport1.Formulas(1) = "CNPJTerceira = '" + Trim(Str(g_Parametros.CNPJ_Terceira)) + "'"
   CrystalReport1.SelectionFormula = sFiltroCheque
   CrystalReport1.Action = 0
   
Else
   
   If sFiltroBordero <> "" Then
      sFiltroBordero = sFiltroBordero + " and {Bordero.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
   Else
      sFiltroBordero = "{Bordero.DataProcessamento}=" + Trim(Str(Geral.DataProcessamento))
   End If
   
   CrystalReport1.ReportFileName = App.path + "\Reports\RelConsultaChequesBordero.rpt"
   CrystalReport1.WindowTitle = "Consulta de Cheques Por Borderô"
   CrystalReport1.Formulas(0) = "DataProcessamento = '" & FormataData(Geral.DataProcessamento, DD_MM_AAAA) & "'"
   CrystalReport1.Formulas(1) = "CNPJTerceira = '" + Trim(Str(g_Parametros.CNPJ_Terceira)) + "'"
   CrystalReport1.SelectionFormula = sFiltroBordero
   CrystalReport1.Action = 0
   
End If

Screen.MousePointer = Default

End Sub

Private Sub cmdSair_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         * Encerra Consulta de Cheque e Borderôs *                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub

Private Sub DteDtaDeposito_GotFocus()

sFiltroCheque = ""

End Sub

Private Sub DteDtaDeposito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdConfirmar.SetFocus
        ''''''''''''''''''''''''
        '* Se data for válida *'
        ''''''''''''''''''''''''
        If Len(Trim(DteDtaDeposito.Text)) <> 0 Then
            Call DteDtaDeposito_LostFocus
            Call cmdConfirmar_Click
        Else
            MsgBox "Informe a Data de Depósito.", vbInformation + vbOKOnly, App.Title
            DteDtaDeposito.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub DteDtaDeposito_LostFocus()

sFiltroCheque = "{Cheque.DataDeposito}=" + Right(DteDtaDeposito.Text, 4) + Mid(DteDtaDeposito.Text, 3, 2) + Left(DteDtaDeposito.Text, 2)

End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                * Carga de Objetos e Formatações *                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    ''''''''''''''''''''''''''''''''''''''''
    ' * Formatação da Data Processamento * '
    ''''''''''''''''''''''''''''''''''''''''
    Call FormataDataProcessamento

    ''''''''''''''''''''''''''''''''''''''''
    ' * Lista Todos os status do Borderô  * '
    ''''''''''''''''''''''''''''''''''''''''
    Call ListaStatusBordero
    
    '''''''''''''''''''''''''''''
    ' * Formatação de Cheques * '
    '''''''''''''''''''''''''''''
    With GrdCheque
        .Cols = 9
        
        ''''''''''''''''''''''''''''''''''
        '   * Nome das Colunas Fixas *   '
        ''''''''''''''''''''''''''''''''''
        .TextMatrix(0, 0) = "Banco"
        .TextMatrix(0, 1) = "Agência"
        .TextMatrix(0, 2) = "Conta Corrente"
        .TextMatrix(0, 3) = "Nr. Cheque"
        .TextMatrix(0, 4) = "Data Depósito"
        .TextMatrix(0, 5) = String(1, " ") & "CNPJ/CPF"
        .TextMatrix(0, 6) = "Situação"
        .TextMatrix(0, 7) = String(15, " ") & "Valor"
        .TextMatrix(0, 8) = "IdBordero"
        
        ''''''''''''''''''''''''''''''''''
        '   * Tamanho de Cada Coluna *   '
        ''''''''''''''''''''''''''''''''''
        .ColWidth(0) = .Width * 0.07
        .ColWidth(1) = .Width * 0.08
        .ColWidth(2) = .Width * 0.12
        .ColWidth(3) = .Width * 0.1
        .ColWidth(4) = .Width * 0.13
        .ColWidth(5) = .Width * 0.19
        .ColWidth(6) = .Width * 0.16
        .ColWidth(7) = .Width * 0.13    'Ajustar largura também na função (PesquisaCheque)
        .ColWidth(8) = .Width * 0.13
        
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

    End With
    
    '''''''''''''''''''''''''''''''''''''''
    ' * Formatação da Grade de Borderôs * '
    '''''''''''''''''''''''''''''''''''''''
    With GrdBordero
        .Cols = 9
        
        ''''''''''''''''''''''''''''''''''
        '   * Nome das Colunas Fixas *   '
        ''''''''''''''''''''''''''''''''''
        .TextMatrix(0, 0) = "Num. Borderô"
        .TextMatrix(0, 1) = "Agência"
        .TextMatrix(0, 2) = "Conta Corrente"
        .TextMatrix(0, 3) = "Tipo Custodia"
        .TextMatrix(0, 4) = "Cod. Loja"
        .TextMatrix(0, 5) = "Data Entrada"
        .TextMatrix(0, 6) = "Nome Cliente"
        .TextMatrix(0, 7) = "IdBordero"
        .TextMatrix(0, 8) = "StatusBordero"
        
        ''''''''''''''''''''''''''''''''''
        '   * Tamanho de Cada Coluna *   '
        ''''''''''''''''''''''''''''''''''
        .ColWidth(0) = .Width * 0.2
        .ColWidth(1) = .Width * 0.08
        .ColWidth(2) = .Width * 0.15
        .ColWidth(3) = .Width * 0.25
        .ColWidth(4) = .Width * 0.15
        .ColWidth(5) = .Width * 0.14
        .ColWidth(6) = .Width * 0.14
        .ColWidth(7) = 0
        .ColWidth(8) = 0

        '''''''''''''''''''''''''''''''''
        '*Alinhamento das Colunas Fixas*'
        '''''''''''''''''''''''''''''''''
        .ColAlignment(0) = 3
        .ColAlignment(1) = 3
        .ColAlignment(2) = 3
        .ColAlignment(3) = 3
        .ColAlignment(4) = 3
        .ColAlignment(5) = 3
    End With
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Formatação da Grade de Datas de Depósito * '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    With GrdDatasDeposito
        .Cols = 3
    
        ''''''''''''''''''''''''''''''''''
        '   * Nome das Colunas Fixas *   '
        ''''''''''''''''''''''''''''''''''
        .TextMatrix(0, 0) = "Data Depósito"
        .TextMatrix(0, 1) = "Quantidade de Cheques"
        .TextMatrix(0, 2) = "Valor do Deposito"
       
        ''''''''''''''''''''''''''''''''''
        '   * Tamanho de Cada Coluna *   '
        ''''''''''''''''''''''''''''''''''
        .ColWidth(0) = .Width * 0.33
        .ColWidth(1) = .Width * 0.33
        .ColWidth(2) = .Width * 0.34
        
        '''''''''''''''''''''''''''''''''
        '*Alinhamento das Colunas Fixas*'
        '''''''''''''''''''''''''''''''''
        .ColAlignment(0) = 3
        .ColAlignment(1) = 3
        .ColAlignment(2) = 3
    End With
    
    
    '''''''''''''''''''''''''''''''''''
    ' * Habilita/Desabilita Tabbled * '
    '''''''''''''''''''''''''''''''''''
    With TabOpcoes
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
    End With
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao Inicializar Formulário.", Err)
    Unload Me
    
End Sub
Private Sub GrdBordero_DblClick()
    ''''''''''''''''''''''''''''''''''''''''''''
    ' * Habilita Tab com Cheques do Borderô  * '
    ''''''''''''''''''''''''''''''''''''''''''''
    txtnumBordero.Text = GrdBordero.TextMatrix(GrdBordero.Row, 0)
    Call PesquisaChequesBordero

End Sub
Private Sub GrdCheque_DblClick()
    ''''''''''''''''''''''''''''''''''''''''''''
    ' * Habilita Tab com Borderô dos Cheques * '
    ''''''''''''''''''''''''''''''''''''''''''''
    Call PesquisaIdBordero
    
End Sub
Private Sub GrdCheque_SelChange()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     * Seleciona Linha Inteira *'                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static m_SelChange As Boolean
    
    'Controle de Acesso
    If m_SelChange = True Then Exit Sub
    
    m_SelChange = True
        
        GrdCheque.Row = GrdCheque.RowSel
        GrdCheque.Col = 0
        GrdCheque.ColSel = 7
        GrdCheque.SetFocus
        
    m_SelChange = False

End Sub
Private Sub GrdBordero_SelChange()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     * Seleciona Linha Inteira *'                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static m_SelChange As Boolean
    
    'Controle de Acesso
    If m_SelChange = True Then Exit Sub
    
    m_SelChange = True
        
        GrdBordero.Row = GrdBordero.RowSel
        Call MontaGradeDeDatas
        GrdBordero.Col = 0
        GrdBordero.ColSel = 7
        GrdBordero.SetFocus
        
    m_SelChange = False

End Sub
Private Sub optbordero_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
        
    optcheque.Value = False
    txtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    optdeposito.Value = False
    DteDtaDeposito.Enabled = False
    optfavorecido.Value = False
    txtfagencia.Enabled = False
    Txtfcc.Enabled = False
    optstatus.Value = False
    CboStatus.Enabled = False
    OptValor.Value = False
    txtValor.Enabled = False

    ''''''''''''''''''''''''''''''
    '    * Habilita Borderô *    '
    ''''''''''''''''''''''''''''''
    txtnumBordero.Enabled = True
        
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me
    
    ''''''''''''''''''''''''''''
    ' * Focaliza Num Borderô * '
    ''''''''''''''''''''''''''''
    txtnumBordero.SetFocus
    
End Sub
Private Sub optcheque_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
        
    optbordero.Value = False
    txtnumBordero.Enabled = False
    optfavorecido.Value = False
    txtfagencia.Enabled = False
    Txtfcc.Enabled = False
    optdeposito.Value = False
    DteDtaDeposito.Enabled = False
    optstatus.Value = False
    CboStatus.Enabled = False
    OptValor.Value = False
    txtValor.Enabled = False

    ''''''''''''''''''''''''''''''
    '     * Habilita Cheque *    '
    ''''''''''''''''''''''''''''''
    txtcBco.Enabled = True
    txtcAgencia.Enabled = True
    TxtCcc.Enabled = True
    
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me

    ''''''''''''''''''''''''''''''
    '    * Focaliza Agencia *    '
    ''''''''''''''''''''''''''''''
    txtcBco.SetFocus

End Sub
Private Sub optdeposito_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
        
    optbordero.Value = False
    txtnumBordero.Enabled = False
    optcheque.Value = False
    txtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    optfavorecido.Value = False
    txtfagencia.Enabled = False
    Txtfcc.Enabled = False
    optstatus.Value = False
    CboStatus.Enabled = False
    OptValor.Value = False
    txtValor.Enabled = False

    ''''''''''''''''''''''''''''''
    ' * Habilita Data Depósito * '
    ''''''''''''''''''''''''''''''
    DteDtaDeposito.Enabled = True
    
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me

    '''''''''''''''''''''''''''''''
    ' * Focaliza Data de Depósito *
    '''''''''''''''''''''''''''''''
    DteDtaDeposito.SetFocus
    
End Sub
Private Sub optfavorecido_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
    
    optbordero.Value = False
    txtnumBordero.Enabled = False
    optcheque.Value = False
    txtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    optdeposito.Value = False
    DteDtaDeposito.Enabled = False
    optstatus.Value = False
    CboStatus.Enabled = False
    OptValor.Value = False
    txtValor.Enabled = False
    
    ''''''''''''''''''''''''''''''
    '   * Habilita Favorecido *  '
    ''''''''''''''''''''''''''''''
    txtfagencia.Enabled = True
    Txtfcc.Enabled = True
    
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me
    
    ''''''''''''''''''''''''''''''
    '    * Focaliza Agência *    '
    ''''''''''''''''''''''''''''''
    txtfagencia.SetFocus
    
End Sub
Private Sub optstatus_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
    
    optbordero.Value = False
    txtnumBordero.Enabled = False
    optcheque.Value = False
    txtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False
    optfavorecido.Value = False
    txtfagencia.Enabled = False
    Txtfcc.Enabled = False
    optdeposito.Value = False
    DteDtaDeposito.Enabled = False
    OptValor.Value = False
    txtValor.Enabled = False

    ''''''''''''''''''''''''''''''
    '     * Habilita Status *    '
    ''''''''''''''''''''''''''''''
    CboStatus.Enabled = True
    
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me
    
    ''''''''''''''''''''''''''''''
    '     * Focaliza Status *    '
    ''''''''''''''''''''''''''''''
    CboStatus.SetFocus
    
End Sub
Private Sub OptValor_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Quando um item for selecionado desmarca os demais *                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ''''''''''''''''''''''''
    '   * Default COMBO *  '
    ''''''''''''''''''''''''
    CboStatus.Text = sStatusDefault
        
    optbordero.Value = False
    txtnumBordero.Enabled = False
    optfavorecido.Value = False
    txtfagencia.Enabled = False
    Txtfcc.Enabled = False
    optdeposito.Value = False
    DteDtaDeposito.Enabled = False
    optstatus.Value = False
    CboStatus.Enabled = False
    optcheque.Value = False
    txtcBco.Enabled = False
    txtcAgencia.Enabled = False
    TxtCcc.Enabled = False

    ''''''''''''''''''''''''''''''
    '     * Habilita Valor  *    '
    ''''''''''''''''''''''''''''''
    txtValor.Enabled = True
    
    ''''''''''''''''''''''''''''''
    ' * Faz Limpeza dos Campos * '
    ''''''''''''''''''''''''''''''
    LimpaTela Me

    ''''''''''''''''''''''''''''''
    '    * Focaliza Agencia *    '
    ''''''''''''''''''''''''''''''
    txtValor.SetFocus

End Sub
Private Sub TabOpcoes_Click(PreviousTab As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       *  Se clicar na Tab 0 desabilita as demais *                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If TabOpcoes.Tab = 0 Then
        
        TabOpcoes.TabEnabled(0) = True
        TabOpcoes.TabEnabled(1) = False
        TabOpcoes.TabEnabled(2) = False
    
        ''''''''''''''''''''''''
        ' * Pesquisa Default * '
        ''''''''''''''''''''''''
        Call optfavorecido_Click
        
        ''''''''''''''''''''''''
        '   * Default COMBO *  '
        ''''''''''''''''''''''''
        CboStatus.Text = sStatusDefault
        
        cmdPrint.Enabled = False
        
    End If
    
    If TabOpcoes.Tab = 2 Then
        Call GrdBordero_SelChange
    End If

End Sub
Private Sub txtcAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtcAgencia = (Format(txtcAgencia, String(4, "0")))
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
End Sub

Private Sub txtcAgencia_LostFocus()

If sFiltroCheque = "" Then
   sFiltroCheque = "Mid({Cheque.CMC7},4,4)='" + Trim(txtcAgencia.Text) + "'"
Else
   sFiltroCheque = sFiltroCheque + " and Mid({Cheque.CMC7},4,4)='" + Trim(txtcAgencia.Text) + "'"
End If

End Sub

Private Sub txtcBco_GotFocus()

sFiltroCheque = ""

End Sub

Private Sub TxtcBco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
End Sub

Private Sub txtcBco_LostFocus()

If sFiltroCheque = "" Then
   sFiltroCheque = "Mid({Cheque.CMC7},1,3)='" + Trim(txtcBco.Text) + "'"
Else
   sFiltroCheque = sFiltroCheque + " Mid({Cheque.CMC7},1,3)='" + Trim(txtcBco.Text) + "'"
End If

End Sub

Private Sub TxtCcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        TxtCcc = (Format(TxtCcc, String(10, "0")))
        TxtCcc_LostFocus
        cmdConfirmar_Click
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
    
End Sub

Private Sub TxtCcc_LostFocus()

If sFiltroCheque = "" Then
   sFiltroCheque = "Mid({Cheque.CMC7},20,10)='" + Trim(TxtCcc.Text) + "'"
Else
   sFiltroCheque = sFiltroCheque + " and Mid({Cheque.CMC7},20,10)='" + Trim(TxtCcc.Text) + "'"
End If

End Sub

Private Sub txtfagencia_GotFocus()

sFiltroBordero = ""

End Sub

Private Sub txtfagencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtfagencia = (Format(txtfagencia, String(4, "0")))
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
End Sub

Private Sub txtfagencia_LostFocus()

If sFiltroBordero = "" Then
   sFiltroBordero = "{Bordero.Agencia}=" + Trim(txtfagencia.Text)
Else
   sFiltroBordero = sFiltroBordero + "{Bordero.Agencia}=" + Trim(txtfagencia.Text)
End If

End Sub

Private Sub Txtfcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtfcc_LostFocus
        cmdConfirmar_Click
        Exit Sub
    End If

    If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
End Sub

Private Sub Txtfcc_LostFocus()

If sFiltroBordero = "" Then
   sFiltroBordero = "{Bordero.Conta}=" + Trim(Txtfcc.Text)
Else
   sFiltroBordero = sFiltroBordero + " and {Bordero.Conta}=" + Trim(Txtfcc.Text)
End If

End Sub

Private Sub txtnumBordero_GotFocus()

sFiltroBordero = ""

End Sub

Private Sub txtnumBordero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtnumBordero.Text = Format(CStr(txtnumBordero.Text), String(19, "0"))
        txtnumBordero_LostFocus
        cmdConfirmar_Click
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
    End If
End Sub

Private Sub txtnumBordero_LostFocus()

sFiltroBordero = "{Bordero.Num_Bordero}='" + Trim(txtnumBordero.Text) + "'"

End Sub

Private Sub txtValor_GotFocus()

sFiltroCheque = ""

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValor_LostFocus
        cmdConfirmar_Click
        Exit Sub
    End If
End Sub

Private Sub txtValor_LostFocus()

sFiltroCheque = "{Cheque.Valor}=" + Trim(txtValor.Text)

End Sub
