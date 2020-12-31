VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Object = "{ED123F48-E23F-11D4-B08D-00600899AB13}#1.0#0"; "UbbEdit.ocx"
Begin VB.Form Bordero 
   Caption         =   "Sistema de Captura - Borderô"
   ClientHeight    =   7212
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   9948
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7212
   ScaleWidth      =   9948
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSomatoriaControle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   288
      TabIndex        =   13
      Top             =   6744
      Width           =   3132
   End
   Begin VB.Frame Frame6 
      Height          =   1188
      Left            =   7728
      TabIndex        =   31
      Top             =   3450
      Width           =   1908
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   345
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1716
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Ca&ncelar"
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1716
      End
      Begin VB.CommandButton cmdProvaZero 
         Caption         =   "Enviar &Prova Zero"
         Height          =   348
         Left            =   96
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1716
      End
   End
   Begin VB.TextBox txtSomatoriaDepositos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   6528
      TabIndex        =   12
      Top             =   6216
      Width           =   3156
   End
   Begin VB.TextBox txtSomatoriaQuantidades 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3408
      TabIndex        =   11
      Top             =   6216
      Width           =   3132
   End
   Begin VB.TextBox txtSomatoriaDatas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   288
      TabIndex        =   10
      Top             =   6216
      Width           =   3132
   End
   Begin VB.Frame Frame5 
      Caption         =   "Nome do Cliente"
      Height          =   588
      Left            =   288
      TabIndex        =   30
      Top             =   1248
      Width           =   9348
      Begin VB.TextBox txtNomeCliente 
         Height          =   315
         Left            =   72
         MaxLength       =   50
         TabIndex        =   3
         Top             =   192
         Width           =   9180
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Número do Borderô"
      Height          =   588
      Left            =   288
      TabIndex        =   29
      Top             =   24
      Width           =   2268
      Begin VB.TextBox txtNumeroBordero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   72
         MaxLength       =   19
         TabIndex        =   0
         Top             =   192
         Width           =   2124
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data de Entrada"
      Height          =   588
      Left            =   7728
      TabIndex        =   28
      Top             =   2712
      Width           =   1908
      Begin DATEEDITLib.DateEdit txtDataEntradaBordero 
         Height          =   324
         Left            =   72
         TabIndex        =   6
         Top             =   216
         Width           =   1740
         _Version        =   65537
         _ExtentX        =   3069
         _ExtentY        =   572
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame fraLoja 
      Caption         =   "Loja"
      Height          =   588
      Left            =   7728
      TabIndex        =   27
      Top             =   1992
      Width           =   1908
      Begin VB.TextBox txtLoja 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   72
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "0"
         Top             =   192
         Width           =   1740
      End
   End
   Begin UbbEdt.UbbEdit txtAgencia 
      Height          =   576
      Left            =   288
      TabIndex        =   1
      Top             =   636
      Width           =   744
      _ExtentX        =   1312
      _ExtentY        =   1016
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   4
      TextMaxNumChars =   4
      Title           =   "Agência"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Custódia"
      Height          =   1476
      Left            =   288
      TabIndex        =   18
      Top             =   1896
      Width           =   4020
      Begin VB.ListBox lstTipoCustodia 
         Height          =   912
         Left            =   72
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   264
         Width           =   3852
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Relação de Datas"
      Height          =   2292
      Left            =   288
      TabIndex        =   19
      Top             =   3456
      Width           =   5460
      Begin VB.CommandButton cmdExcluirData 
         Caption         =   "Excluir"
         Height          =   300
         Left            =   4392
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1176
         Width           =   924
      End
      Begin VB.CommandButton cmdInserirData 
         Caption         =   "Inserir"
         Height          =   300
         Left            =   4392
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   816
         Width           =   924
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   8
         Top             =   480
         Width           =   996
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDatas 
         Height          =   1380
         Left            =   72
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   816
         Width           =   4188
         _ExtentX        =   7387
         _ExtentY        =   2434
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         FormatString    =   "Data de Depósito    |Quantidade   |Valor do Depósito  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UbbEdt.UbbEdit txtValorDeposito 
         Height          =   504
         Left            =   2568
         TabIndex        =   9
         Top             =   276
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   889
         TextColor       =   -2147483640
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   2
         TextMaxNumChars =   14
         BorderStyle     =   0
         Title           =   "Valor do Depósito"
         AutoNextControl =   0   'False
      End
      Begin DATEEDITLib.DateEdit txtDataDeposito 
         Height          =   288
         Left            =   72
         TabIndex        =   7
         Top             =   480
         Width           =   1476
         _Version        =   65537
         _ExtentX        =   2603
         _ExtentY        =   508
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         Height          =   192
         Left            =   1560
         TabIndex        =   21
         Top             =   264
         Width           =   840
      End
      Begin VB.Label lblDataDeposito 
         AutoSize        =   -1  'True
         Caption         =   "Data de Depósito"
         Height          =   192
         Left            =   96
         TabIndex        =   20
         Top             =   264
         Width           =   1272
      End
   End
   Begin UbbEdt.UbbEdit txtContaCorrente 
      Height          =   576
      Left            =   1440
      TabIndex        =   2
      Top             =   636
      Width           =   1128
      _ExtentX        =   1990
      _ExtentY        =   1016
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   15
      TextMaxNumChars =   10
      Title           =   "Conta"
   End
   Begin UbbEdt.UBBValid UBBValid1 
      Left            =   4992
      Top             =   516
      _ExtentX        =   635
      _ExtentY        =   656
      Banco           =   409
      ColorOK         =   0
      Campo12         =   "txtAgencia"
      Campo13         =   "txtContaCorrente"
   End
   Begin VB.Label lblSomatoria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   3408
      TabIndex        =   26
      Top             =   6744
      Width           =   6276
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                     Somatória de Controle"
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   288
      TabIndex        =   25
      Top             =   6480
      Width           =   9396
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 - Somatória de Depósitos"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6528
      TabIndex        =   24
      Top             =   5976
      Width           =   3156
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 - Somatória de Quantidades"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3408
      TabIndex        =   23
      Top             =   5976
      Width           =   3132
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 - Somatória de Datas"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   288
      TabIndex        =   22
      Top             =   5976
      Width           =   3132
   End
End
Attribute VB_Name = "Bordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

 '''''''''''''''''''''''''''''''''''''''''''''''''''''
 '            Definição de Tipos Privados
 '''''''''''''''''''''''''''''''''''''''''''''''''''''
 Private Type tpBorderoDatas
     Data_Deposito                               As Long
     Quantidade_Cheques                          As String
     Valor_Deposito                              As String
     Alteracao                                   As Boolean
     Inclusao                                    As Boolean
     Exclusao                                    As Boolean
     Antiga_Data                                 As Long
 End Type

 Private Type tpConsistencia
     sDatasExistentes()                          As tpBorderoDatas
     iQuantidade_Linhas_Grid                     As Integer
     iQuantidade_Cheques_Grid                    As Integer
 End Type

 Private Type tpBordero
     Numero_Bordero                              As String
     Agencia                                     As String
     Conta_Corrente                              As String
     Nome_Cliente                                As String
     Id_Custodia                                 As String
     Id_Bordero                                  As Long
     Loja                                        As String
     Data_Entrada_Bordero                        As Long
     Datas()                                     As tpBorderoDatas
     Somatoria_Datas                             As String
     Somatoria_Quantidades                       As String
     Somatoria_Depositos                         As String
     Somatoria_Controle                          As String
 End Type

 '''''''''''''''''''''''''''''
 'Constantes do grid de Datas'
 '''''''''''''''''''''''''''''
 Private Const COL_GRD_DATADEPOSITO = 0
 Private Const COL_GRD_QUANTIDADE = 1
 Private Const COL_GRD_VALOR = 2

 Private Enum eRetornoGrid
     eGR_ProximoControle = 0
     eGR_Erro = 1
     eGR_Ok
 End Enum
 
 Private Enum eTipoData
     eFeriadoNacional
     eFeriadoLocal
     eFinalDeSemana
     eDiaDeSemana
     eDiasExcedente
 End Enum

 Private Enum eModoEdicao
     eModo_Inclusao
     eModo_Alteracao
 End Enum

 Dim m_fColor                                    As Double
 Dim m_bColor                                    As Double
 Dim m_Consistencia                              As tpConsistencia
 Dim m_IdBordero                                 As Long
 Dim m_IdBorderoIncluso                          As Long
 Dim m_Num_Bordero                               As String
 Dim m_StartarBordero                            As Boolean
 Dim m_Modo                                      As eModoEdicao
 Dim m_RetornoBordero                            As enumRetornoModal
 Dim m_DataProcessamento                         As Long
 Dim m_Event                                     As Boolean
 Dim m_DatasExclusao()                           As tpBorderoDatas

'Vetor p/ Retornar DE datas e qtde de cheques p/ complementacao
 Private VetDatas                                    As Variant
Private Function ExistePendenciaGrid() As Boolean

    Dim i       As Integer

    ExistePendenciaGrid = False
    
    For i = 1 To GrdDatas.Rows - 1
        If GrdDatas.RowData(i) = 1 Then
            ExistePendenciaGrid = True
            Exit For
        End If
    Next i

End Function
Private Sub LimpaGridDatas()

    Dim i       As Integer
    
    '''''''''''''''''''''''''''''''''''''''''
    'Não pode envocar o método Clear do Grid'
    'porque o Clear tira também o Cabecalho '
    '''''''''''''''''''''''''''''''''''''''''
    GrdDatas.Rows = 1
    GrdDatas.BackColorSel = GrdDatas.BackColorFixed
    GrdDatas.ForeColorSel = GrdDatas.ForeColorSel
    

End Sub
Private Sub LimpaPendenciaGrid()

    Dim i       As Integer

    For i = 1 To GrdDatas.Rows - 1
        GrdDatas.RowData(i) = 0
    Next i

End Sub
Private Sub LimpaTelaBordero()


    txtNumeroBordero.Text = ""
    txtAgencia.Text = ""
    txtContaCorrente.Text = ""
    txtNomeCliente.Text = ""
    lstTipoCustodia.Selected(0) = True
    txtLoja.Text = ""
    txtDataEntradaBordero.Text = ""
    
    LimpaCamposData
    LimpaPendenciaGrid
    LimpaGridDatas

    txtSomatoriaDatas.Text = ""
    txtSomatoriaQuantidades.Text = ""
    txtSomatoriaDepositos.Text = ""
    txtSomatoriaControle.Text = ""

End Sub
Private Function MostraBordero(ByVal pIdBordero As Long) As Boolean

    Dim rst                 As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim docBordero          As tpBordero
    
    On Error GoTo Erro_MostraBordero:
    
    MostraBordero = False
    
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero(m_DataProcessamento, pIdBordero))
    
    If rst.EOF() Then Exit Function
    
    ''''''''''''''''''''''''''
    'Carrega dados do Bordero'
    ''''''''''''''''''''''''''
    With docBordero
        .Agencia = rst!Agencia
        .Conta_Corrente = rst!Conta
        .Data_Entrada_Bordero = rst!DataEntrada
        .Id_Custodia = rst!CodigoCarteira
        .Loja = FormataString(rst!CodigoLoja, "0", txtLoja.MaxLength, True)
        .Nome_Cliente = rst!NomeCliente
        .Numero_Bordero = FormataString(rst!Num_Bordero, "0", txtNumeroBordero.MaxLength, True)
        .Somatoria_Datas = rst!SomaData
        .Somatoria_Depositos = rst!SomaValor
        .Somatoria_Quantidades = rst!SomaQuantidade
        .Somatoria_Controle = rst!SomaTodos
        .Id_Bordero = rst!IdBordero
    End With
    
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetDatasBordero(m_DataProcessamento, docBordero.Id_Bordero))
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Preenche Grid de Datas e Vetor de Consistencias'
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Do While Not rst.EOF()
        ReDim Preserve docBordero.Datas(rst.AbsolutePosition - 1) As tpBorderoDatas
        ReDim Preserve m_Consistencia.sDatasExistentes(rst.AbsolutePosition - 1) As tpBorderoDatas
        
        docBordero.Datas(rst.AbsolutePosition - 1).Data_Deposito = rst!DataDeposito
        docBordero.Datas(rst.AbsolutePosition - 1).Antiga_Data = rst!DataDeposito
        docBordero.Datas(rst.AbsolutePosition - 1).Quantidade_Cheques = rst!QuantidadeCheques
        docBordero.Datas(rst.AbsolutePosition - 1).Valor_Deposito = rst!ValorDeposito
        ''''''''''''''''''''''''''
        'loop da relação de Datas'
        ''''''''''''''''''''''''''
        GrdDatas.Rows = rst.AbsolutePosition + 1
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_GRD_DATADEPOSITO) = _
                              Format(Format(docBordero.Datas(rst.AbsolutePosition - 1).Data_Deposito, "0000/00/00"), "dd/mm/yyyy")
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_GRD_QUANTIDADE) = _
                              docBordero.Datas(rst.AbsolutePosition - 1).Quantidade_Cheques
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_GRD_VALOR) = _
                              Format(docBordero.Datas(rst.AbsolutePosition - 1).Valor_Deposito, MASK_VALOR)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Preenche o vetor de consistencia da Data de Deposito'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_Consistencia.sDatasExistentes(rst.AbsolutePosition - 1).Data_Deposito = _
                                        docBordero.Datas(rst.AbsolutePosition - 1).Data_Deposito
        m_Consistencia.sDatasExistentes(rst.AbsolutePosition - 1).Antiga_Data = _
                                        docBordero.Datas(rst.AbsolutePosition - 1).Antiga_Data
        m_Consistencia.iQuantidade_Cheques_Grid = m_Consistencia.iQuantidade_Cheques_Grid + 1
        ' Alteração Total  Linhas Grig em 29.12.01
        'm_Consistencia.iQuantidade_Linhas_Grid = m_Consistencia.iQuantidade_Cheques_Grid + 1
        m_Consistencia.iQuantidade_Linhas_Grid = m_Consistencia.iQuantidade_Linhas_Grid + 1
        rst.MoveNext
    Loop
        
    rst.Close
    
    ''''''''''''''''''''''''''''''''
    'Preenche os campos necessários'
    ''''''''''''''''''''''''''''''''
    With docBordero
        txtNumeroBordero.Text = .Numero_Bordero
        txtAgencia.Text = .Agencia
        txtContaCorrente.Text = .Conta_Corrente
        txtNomeCliente.Text = .Nome_Cliente
        txtLoja.Text = .Loja
        txtDataEntradaBordero.Text = Format(Format(.Data_Entrada_Bordero, "0000/00/00"), "dd/mm/yyyy")
        txtSomatoriaDatas.Text = .Somatoria_Datas
        txtSomatoriaQuantidades.Text = .Somatoria_Quantidades
        txtSomatoriaDepositos.Text = .Somatoria_Depositos
        txtSomatoriaControle.Text = .Somatoria_Controle
    End With
    '''''''''''''''''''''''''''''''
    'Seleciona a custodia definida'
    '''''''''''''''''''''''''''''''
    SelecionarCustodia docBordero.Id_Custodia
    
    
    MostraBordero = True
    Exit Function
Erro_MostraBordero:

End Function

Private Function ColocaPonto(ByVal pValor As String) As String

    Dim sstr            As String
    
    sstr = Format(pValor, "0000000000000000000000.00")
    
    
    If InStr(sstr, ",") Then
        Mid(sstr, InStr(sstr, ","), 1) = "."
    End If

    ColocaPonto = sstr
End Function

Private Sub SelecionaLista(ByVal pItem As Integer)

    Dim i As Integer
    
    For i = 0 To lstTipoCustodia.ListCount - 1
        lstTipoCustodia.Selected(i) = False
    Next i
    
    lstTipoCustodia.Selected(pItem) = True
    lstTipoCustodia.ListIndex = pItem

End Sub

Private Sub SelecionarCustodia(ByVal pIdCustodia As Integer)

    Dim i As Integer
    
    For i = 0 To lstTipoCustodia.ListCount - 1
        If lstTipoCustodia.ItemData(i) = pIdCustodia Then
            lstTipoCustodia.Selected(i) = True
            Exit For
        End If
    Next i

End Sub

Public Sub SetIdbordero(ByVal pIdBordero As Long)

    If Val(pIdBordero) = 0 Then
        m_IdBordero = 0
        m_IdBorderoIncluso = 0
        Exit Sub
    End If

    m_IdBordero = pIdBordero
    m_StartarBordero = True

End Sub
Public Function ShowModal(Optional ByRef pIdBordero As Long, Optional ByRef pNum_Bordero As String, Optional ByVal pVem_De_Prova_Zero As Boolean = False, Optional ByRef pVetDatas As Variant) As enumRetornoModal
    ''''''''''''''''''''''''''''''''''''''''''''''
    'Habilita ou Desabilita o botão de Prova Zero'
    'quando vem de Prova Zero                    '
    ''''''''''''''''''''''''''''''''''''''''''''''
    cmdProvaZero.Enabled = Not pVem_De_Prova_Zero

    Me.Show vbModal

    pIdBordero = m_IdBorderoIncluso
    pNum_Bordero = m_Num_Bordero
    pVetDatas = VetDatas

    ShowModal = m_RetornoBordero

End Function

Private Function ValidaData(ByVal pData As String, ByVal pAgencia As String) As eTipoData

    Dim rst                     As New ADODB.Recordset
    Dim Proc                    As New Custodia.Selecionar
    Dim sDataProcessamento      As String
    Dim eRetornoData            As eTipoData
    Dim eRetornoData2           As eTipoData
    Dim i                       As Integer
    Dim sData                   As String
    Dim iDiasUteis              As Integer

    On Error GoTo Erro_DataValida
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Converte data de processamento para dd/mm/yyyy'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    sDataProcessamento = FormataData(Geral.DataProcessamento, DD_MM_AAAA)
    'sDataProcessamento = "27/04/2001"
    
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Faz Loop dos dias contando os dias uteis'
    ''''''''''''''''''''''''''''''''''''''''''
    
    sData = Format(DateSerial(Year(sDataProcessamento), Month(sDataProcessamento), Day(sDataProcessamento) + 1), "dd/mm/yyyy")
    
 
    
    '''''''''''''''''''''''''''''''
    'Verifica se é final de semana'
    '''''''''''''''''''''''''''''''
'    If (Weekday(pData, vbSunday) = 1) Or _
'       (Weekday(pData, vbSunday) = 7) Then
'        eRetornoData = eFinalDeSemana
'    Else
                    eRetornoData = eDiaDeSemana
'    End If

    ''''''''''''''''''''''''''''''''''''''''''''''
    'Depois que foi excluída a tabela de feriados'
    'não é mais necessário fazer este loop. Mas  '
    'como não temos tempo...                     '
    ''''''''''''''''''''''''''''''''''''''''''''''

     sData = Format(CLng(Mid(sData, 1, 2) & Mid(sData, 4, 2) & Mid(sData, 7, 4)), "00000000")
     pData = Format(CLng(Mid(pData, 1, 2) & Mid(pData, 4, 2) & Mid(pData, 7, 4)), "00000000")
     
         Do While DataAAAAMMDD(sData) <= DataAAAAMMDD(pData)
        '''''''''''''''''''''''''''''''
        'Verifica se é final de semana'
        '''''''''''''''''''''''''''''''
        sData = Mid(sData, 1, 2) & "/" & Mid(sData, 3, 2) & "/" & Mid(sData, 5, 4)
        If (Weekday(sData, vbSunday) = 1) Or _
           (Weekday(sData, vbSunday) = 7) Then

            'eRetornoData = eFinalDeSemana

        Else
            iDiasUteis = iDiasUteis + 1
            eRetornoData = eDiaDeSemana
'            '''''''''''''''''''''''''''''''
'            'Consulta a Tabela de Feriados'
'            '''''''''''''''''''''''''''''''
'            Set rst = g_cMainConnection.Execute(Proc.GetFeriado(Format(sData, "mm/dd/yyyy"), pAgencia))
'
'            If Not rst.EOF() Then
'                If rst!TipoFeriado = 8 Then
'                    eRetornoData = eFeriadoLocal
'                ElseIf rst!TipoFeriado = 9 Then
'                    eRetornoData = eFeriadoNacional
'                Else
'                    iDiasUteis = iDiasUteis + 1
'                End If
'            Else
'                iDiasUteis = iDiasUteis + 1
'                eRetornoData = eDiaDeSemana
'            End If
'            rst.Close

            End If
        'sData = DateSerial(Year(sData), Month(sData), Day(sData) + 1)
        sData = Format(DateSerial(Year(sData), Month(sData), Day(sData) + 1), "dd/mm/yyyy")
        sData = Format(CLng(Mid(sData, 1, 2) & Mid(sData, 4, 2) & Mid(sData, 7, 4)), "00000000")
    Loop

    If eRetornoData = eDiaDeSemana Then
        'If iDiasUteis < 3 Then
        If iDiasUteis < g_Parametros.QuantidadeMinimaDias Then
            eRetornoData = eDiasExcedente
        End If
    End If
    
    ValidaData = eRetornoData
 
    Exit Function
    
Erro_DataValida:

    End Function
'
'
'           Exlui Data do Grid
'
'
Private Sub ExcluirData()

    Dim i                   As Integer
    Dim sDataDeposito       As String
    Dim lRetorno            As Long
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim rst                 As New ADODB.Recordset
    
    On Error GoTo Erro_ExcluiData
    
    ''''''''''''''''''''''''
    'Não pode ser a linha 0'
    ''''''''''''''''''''''''
    If GrdDatas.Rows > 1 Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Verifica se a data à ser excluida possui cheques'
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        If m_Modo = eModo_Alteracao Then
        
            sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_DATADEPOSITO), "yyyymmdd")

            Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetChequesBordero( _
                                                m_DataProcessamento, _
                                                m_IdBordero, _
                                                sDataDeposito), _
                                        lRetorno, _
                                        adCmdText)
            If Not rst.EOF <> 0 Then
                If MsgBox("Confirma a exclusão desta Data e de todos os cheques relacionados.", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
                    Exit Sub
                End If
            End If
            
            On Error Resume Next
            i = UBound(m_DatasExclusao)
            
            If Err <> 0 Then
                ReDim m_DatasExclusao(0) As tpBorderoDatas
            Else
                ReDim Preserve m_DatasExclusao(i) As tpBorderoDatas
            End If
            
            On Error GoTo Erro_ExcluiData
            
            m_DatasExclusao(i).Data_Deposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_DATADEPOSITO), "yyyymmdd")
            m_DatasExclusao(i).Quantidade_Cheques = GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_QUANTIDADE)
            m_DatasExclusao(i).Valor_Deposito = GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_VALOR)
            m_DatasExclusao(i).Exclusao = True
            
        End If
    
        GrdDatas.BackColorSel = GrdDatas.BackColor
        GrdDatas.ForeColorSel = GrdDatas.ForeColor
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Obtem a linha e quantidade de cheques para decrementar no total de cheques'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_Consistencia.iQuantidade_Cheques_Grid = m_Consistencia.iQuantidade_Cheques_Grid - _
                                                  Val(GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_QUANTIDADE))

        '''''''''''''''''''''''''''''''''''''''''''''''''
        'Obtem a data à ser excluída da relação de datas'
        '''''''''''''''''''''''''''''''''''''''''''''''''
        sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_DATADEPOSITO), "yyyymmdd")
        For i = 0 To UBound(m_Consistencia.sDatasExistentes)
            If sDataDeposito = m_Consistencia.sDatasExistentes(i).Data_Deposito Then
                m_Consistencia.sDatasExistentes(i).Data_Deposito = 0
                Exit For
            End If
        Next i
        '''''''''''''''''''''''''''''''''''''''''''
        'Decrementa uma linha na relação de linhas'
        '''''''''''''''''''''''''''''''''''''''''''
        m_Consistencia.iQuantidade_Linhas_Grid = m_Consistencia.iQuantidade_Linhas_Grid - 1

        If GrdDatas.Rows = 2 Then
            GrdDatas.Rows = 1
        Else
            ''''''''''''''''''''''''
            'Remove a linha do Grid'
            ''''''''''''''''''''''''
            GrdDatas.RemoveItem GrdDatas.Row
        End If
        txtDataDeposito.SetFocus
        ''''''''''''''''''''''''''''''''''''''''''''''''
        'Se sobrou somente uma linha que é a linha fixa'
        'então a cor de seleção será a mesma que a cor '
        'da linha fixa                                 '
        ''''''''''''''''''''''''''''''''''''''''''''''''
        If GrdDatas.Rows = 1 Then
            GrdDatas.BackColorSel = GrdDatas.BackColorFixed
            GrdDatas.ForeColorSel = GrdDatas.ForeColorFixed
            txtDataDeposito.SetFocus
        End If
    End If
    LimpaCamposData
    
    Exit Sub
    
Erro_ExcluiData:

    TratamentoErro "Erro ao excluir a data deste Borderô.", Err
End Sub

Private Function IncluiNoGrid() As eRetornoGrid

    Dim iLineIndex          As Long
    Dim i                   As Integer
    Dim bRedim              As Boolean

    On Error GoTo Erro_IncluiNoGrid
    
    bRedim = True
    
    ''''''''''''''''''''''''''''''''''
    'Se tiver tudo em branco sai fora'
    'para não inserir no Grid        '
    ''''''''''''''''''''''''''''''''''
    If Trim(txtDataDeposito.Text) = "" And _
       Trim(txtQuantidade.Text) = "" And _
       Trim(txtValorDeposito.Text) = "" Then
       
        IncluiNoGrid = eGR_ProximoControle
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Localiza no Grid se existe alguma linha à ser alterada'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To GrdDatas.Rows - 1
        If Val(GrdDatas.RowData(i)) = 1 Then
            bRedim = False
            iLineIndex = i
            ''''''''''''''''''''''''''''''''''''''''''''
            'FLAG para não alterar esta linha novamente'
            ''''''''''''''''''''''''''''''''''''''''''''
            GrdDatas.RowData(i) = 0
            
            m_Consistencia.sDatasExistentes(i - 1).Data_Deposito = Format(txtDataDeposito.MaskText, "yyyymmdd")
            m_Consistencia.sDatasExistentes(i - 1).Alteracao = True
            'm_Consistencia.sDatasExistentes(i - 1).Inclusao = False
            Exit For
        End If
    Next i


    ''''''''''''''''''''''''''''''''''''''''''''''
    'Verifico se iLineIndex = 0, se for          '
    'defino iLineIndex = Ultima linha porque     '
    'será processado uma inclusao e não alteração'
    ''''''''''''''''''''''''''''''''''''''''''''''
    If iLineIndex = 0 Then
        '''''''''''''''''''''''''''''''''''''''
        'Define o Numero de linhas necessárias'
        '''''''''''''''''''''''''''''''''''''''
        GrdDatas.Rows = GrdDatas.Rows + 1
    
        ''''''''''''''''''''''''''''
        'Sempre será a ultima linha'
        ''''''''''''''''''''''''''''
        iLineIndex = GrdDatas.Rows - 1
        
        '''''''''''''''''''''''''''''''''''''''''''''''''
        'Incrementa quantidade de linhas na consistencia'
        '''''''''''''''''''''''''''''''''''''''''''''''''
        m_Consistencia.iQuantidade_Linhas_Grid = m_Consistencia.iQuantidade_Linhas_Grid + 1
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Incrementa quantidade de cheques inseridas no Grid'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_Consistencia.iQuantidade_Cheques_Grid = m_Consistencia.iQuantidade_Cheques_Grid + Val(txtQuantidade.Text)
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Neste momento já foi adicionada uma linha no Grid'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    GrdDatas.TextMatrix(iLineIndex, COL_GRD_DATADEPOSITO) = txtDataDeposito.MaskText
    GrdDatas.TextMatrix(iLineIndex, COL_GRD_QUANTIDADE) = txtQuantidade.Text
    GrdDatas.TextMatrix(iLineIndex, COL_GRD_VALOR) = Format(Val(InserePonto(txtValorDeposito.Text)), MASK_VALOR)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Inclui a data no Array, no indice está -2 porque sempre no grid terá uma linha'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If bRedim Then
        ReDim Preserve m_Consistencia.sDatasExistentes(GrdDatas.Rows - 2) As tpBorderoDatas
        m_Consistencia.sDatasExistentes(GrdDatas.Rows - 2).Data_Deposito = Format(txtDataDeposito.MaskText, "yyyymmdd")
        m_Consistencia.sDatasExistentes(GrdDatas.Rows - 2).Inclusao = True
    End If
    '''''''''''''''''''''''''''
    'Torna a retirar a seleção'
    '''''''''''''''''''''''''''
    GrdDatas.BackColorSel = GrdDatas.BackColor
    GrdDatas.ForeColorSel = GrdDatas.ForeColor
    
'    '''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Faz a somatória das datas e joga no grid de baixo'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''
'    For i = 0 To UBound(m_Consistencia.sDatasExistentes)
'        If Trim(m_Consistencia.sDatasExistentes(i)) <> "" Then
'
'        End If
'    Next i
    
    IncluiNoGrid = eGR_Ok
    
    Exit Function
Erro_IncluiNoGrid:

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se ocorreu algum erro remove a linha que fora adicionada'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'isto é porque o grid não deixa remover a unica linha não fixa do grid
    If GrdDatas.Rows = 2 Then
        GrdDatas.Rows = 1
    Else
        GrdDatas.RemoveItem iLineIndex
    End If
    GrdDatas.BackColorSel = GrdDatas.BackColorFixed
    GrdDatas.ForeColorSel = GrdDatas.ForeColorFixed
    
    
    IncluiNoGrid = eGR_Erro

End Function

Private Sub LimpaCamposData()
    
    txtDataDeposito.Text = ""
    txtQuantidade.Text = ""
    txtValorDeposito.Text = ""
    
    
    
End Sub
Private Function ValidaCamposData() As Boolean

    Dim i               As Integer
    Dim eRetornoData    As eTipoData
    Dim sstr            As String
    Dim TotalQtdeDatas  As Double
    
    ValidaCamposData = False
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Valida quantidade de itens ja inseridos no grid. Não pode ser mais que no parametro'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not ExistePendenciaGrid() Then
        If (m_Consistencia.iQuantidade_Linhas_Grid) >= Val(g_Parametros.QuantidadeDatas) Then
            MsgBox "A quantidade de datas excedeu o limite de " & g_Parametros.QuantidadeDatas & ".", vbExclamation, Me.Caption
            Exit Function
        End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Não pode exceder o limite de cheques e na inserção pelo menos um cheque'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Val(txtQuantidade.Text) < 1 And Trim(txtQuantidade.Text) <> "") Then
        MsgBox "A quantidade de cheques é inválida.", vbExclamation, Me.Caption
        Exit Function
    End If
        
    ''''''''''''''''''''''''''''''''''''''''''
    'Verifica se já não existe a data no Grid'
    ''''''''''''''''''''''''''''''''''''''''''
    
    If GrdDatas.Rows > 1 Then
        For i = 1 To GrdDatas.Rows - 1
            If m_Consistencia.sDatasExistentes(i - 1).Data_Deposito = Format(Trim(txtDataDeposito.MaskText), "yyyymmdd") And _
               m_Consistencia.sDatasExistentes(i - 1).Data_Deposito <> 0 And _
               GrdDatas.RowData(i) = 0 Then
                MsgBox "Não é permitido a inclusão de datas repetidas na lista.", vbExclamation, Me.Caption
                Exit Function
            End If
            
           'Acumula Qtdes, exceto linha selecionada (ou seja escolhida para alteração)
            TotalQtdeDatas = TotalQtdeDatas + IIf(GrdDatas.RowSel = i And GrdDatas.BackColorSel = &H8000000D, 0, Val(GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE)))
        Next i
    End If
    
   'Soma o valor Digitado para inclusao
    TotalQtdeDatas = TotalQtdeDatas + Val(txtQuantidade.Text)
    
   'Verifica se o total digitado não ultrapassou o parametro
    If TotalQtdeDatas > g_Parametros.QuantidadeCheques Then
        MsgBox "A quantidade de cheques excedeu o limite de " & g_Parametros.QuantidadeCheques & ", especificado em Parâmetro.", vbExclamation, Me.Caption
        Exit Function
    End If
    
    '''''''''''''''''''''''''''''''''''''''
    'Se algum campo não estiver preenchido'
    'é porque não está válido.            '
    '''''''''''''''''''''''''''''''''''''''
    If (Trim(txtDataDeposito.Text) = "") Or _
       (Trim(txtQuantidade.Text) = "") Or _
       (Trim(txtValorDeposito.Text) = "") Then
       
        '''''''''''''''''''''''''''''''
        'Más se tudo estiver vazio,   '
        'manda para o próximo controle'
        '''''''''''''''''''''''''''''''
        If (Trim(txtDataDeposito.Text) = "") And _
           (Trim(txtQuantidade.Text) = "") And _
           (Trim(txtValorDeposito.Text) = "") Then
            
            ValidaCamposData = True
            Exit Function
        End If
            
        MsgBox "Campos inválidos.", vbExclamation, Me.Caption
        Exit Function
        
    End If
    
    
    If Trim(txtAgencia.Text) = "" Then
        MsgBox "É obrigatório o preenchimento da Agência.", vbExclamation, Me.Caption
        Exit Function
    End If
    
    '''''''''''''''''''''''''''
    'Valida se a data é válida'
    '''''''''''''''''''''''''''
    eRetornoData = ValidaData(txtDataDeposito.MaskText, txtAgencia.Text)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se a data do deposito é menor que a data de entrada '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Trim(txtDataDeposito.InverseText) < Trim(txtDataEntradaBordero.InverseText) Then
        MsgBox "A data de depósito não pode ser menor que a data de entrada.", vbExclamation, Me.Caption
        Exit Function
    End If

    If Len(Trim(txtDataEntradaBordero.Text)) = 0 Then
        MsgBox "É obrigatório o preenchimento da data de entrada antes de informar as datas de depósito.", vbExclamation, Me.Caption
        Exit Function
    
    End If

    If eRetornoData = eDiasExcedente Then
        sstr = "Data de depósito menor do que o permitido."
    ElseIf eRetornoData = eFeriadoLocal Then
        sstr = "Esta Agência se encontra em um Feriado Local."
    ElseIf eRetornoData = eFeriadoNacional Then
        sstr = "Esta Agência se encontra em Feriado Nacional."
    ElseIf eRetornoData = eFinalDeSemana Then
        sstr = "Não é permitido data de depósito em Final de Semana."
    End If

    If Trim(sstr) <> "" Then
        MsgBox sstr, vbExclamation, Me.Caption
        Exit Function
    End If

    ValidaCamposData = True

End Function
Private Sub cmdCancelar_Click()

    m_RetornoBordero = eRetornoCancelar

    Unload Me

End Sub
Private Sub cmdConfirmar_Click()

    Dim i                   As Integer
    Dim j                   As Integer
    Dim lRetorno            As Long
    Dim bSelecionado        As Boolean
    Dim iSomaData           As Double
    Dim iQuantidade         As Integer
    Dim iValorDeposito      As Double
    Dim dSomatoriaGeral     As Double
    Dim sstr                As String
    Dim Proc_Alterar        As New Custodia.Atualizar
    Dim Proc_Inserir        As New Custodia.Inserir
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim Proc_Excluir        As New Custodia.Excluir
    Dim rst                 As New ADODB.Recordset
    
    On Error GoTo Erro_Confirmar:
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Valida número de Borderô, não pode estar zerado'
    '''''''''''''''''''''''''''''''''''''''''''''''''
    If Val(txtNumeroBordero.Text) <= 0 Then
        MsgBox "Número de Borderô está inválido.", vbExclamation, Me.Caption
        txtNumeroBordero.SetFocus
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se já não existe o número de Borderô'
    '''''''''''''''''''''''''''''''''''''''''''''''
    If m_Modo = eModo_Inclusao Then
    
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetNumeroBordero(m_DataProcessamento, txtNumeroBordero.Text))
    
        If Not rst.EOF() Then
            MsgBox "Número de Borderô já existente.", vbExclamation, Me.Caption
            txtNumeroBordero.SetFocus
            Exit Sub
        End If

        rst.Close
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    'Valida se foram digitados Agencia e Conta'
    '''''''''''''''''''''''''''''''''''''''''''
    If (Trim(txtAgencia.Text) = "") Or (txtAgencia.TextColor = UBBValid1.ColorInvalid) Then
        MsgBox "O preenchimento do campo Agência/Conta está inválido.", vbExclamation, Me.Caption
        txtAgencia.SetFocus
        Exit Sub
    End If
    If (Trim(txtContaCorrente.Text) = "") Or (txtContaCorrente.TextColor = UBBValid1.ColorInvalid) Then
        MsgBox "O preenchimento do campo Conta Corrente está inválido.", vbExclamation, Me.Caption
        txtContaCorrente.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''
    'Valida campo Nome do Cliente'
    ''''''''''''''''''''''''''''''
    ' Nome não será mais obrigatório - Fase 2
'    If Trim(txtNomeCliente.Text) = "" Then
'        MsgBox "É obrigatório o preenchimento do campo Nome do Cliente.", vbExclamation, Me.Caption
'        txtNomeCliente.SetFocus
'        Exit Sub
'    End If
    
    '''''''''''''''''''''''''''''''''''
    'Valida Lista de Tipos de Custódia'
    '''''''''''''''''''''''''''''''''''
    For i = 0 To lstTipoCustodia.ListCount - 1
        If lstTipoCustodia.Selected(i) = True Then
            bSelecionado = True
            Exit For
        End If
    Next i

    If bSelecionado = False Then
        MsgBox "É obrigatório selecionar um Tipo de Custódia.", vbExclamation, Me.Caption
        lstTipoCustodia.SetFocus
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    'Valida número da Loja, não pode estar zerado'
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Loja não será criticada
    'If Val(txtLoja.Text) <= 0 Then
    '    MsgBox "Número da Loja está inválido.", vbExclamation, Me.Caption
    '    txtLoja.SetFocus
    '    Exit Sub
    'End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Valida data de entrada do Borderô, não pode estar zerado'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Trim(txtDataEntradaBordero.Text) = "" Then
        MsgBox "Data de Entrada do Borderô está inválido.", vbExclamation, Me.Caption
        txtDataEntradaBordero.SetFocus
        Exit Sub
    End If

    If Trim(txtDataEntradaBordero.InverseText) < Format(Now, "yyyymmdd") Then
        MsgBox "Data de Entrada do Borderô não pode ser menor que a data do sistema.", vbExclamation, Me.Caption
        txtDataEntradaBordero.SetFocus
        Exit Sub
    End If

    ''''''''''''''''''''''
    'Valida Grid de Datas'
    ''''''''''''''''''''''
    If GrdDatas.Rows < 2 Then
        MsgBox "É obrigatório o preenchimento da lista de Datas.", vbExclamation, Me.Caption
        txtDataDeposito.SetFocus
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''
    'Valida campo Somatoria de Datas'
    '''''''''''''''''''''''''''''''''
    ' Somatória de Datas automático com parametro
    If g_Parametros.chkSoma Then   ' Se Critica Soma
        
        If Trim(txtSomatoriaDatas.Text) = "" Then
            MsgBox "É obrigatório o preenchimento do campo Somatória de Datas.", vbExclamation, Me.Caption
            txtSomatoriaDatas.SetFocus
            Exit Sub
        End If
            
    End If

    '''''''''''''''''''''''''''''''''''''''
    'Valida campo Somatoria de Quantidades'
    '''''''''''''''''''''''''''''''''''''''
    ' Somatória de Datas automático com parametro
    If g_Parametros.chkSoma Then   ' Se Critica Soma
    
        If Trim(txtSomatoriaQuantidades.Text) = "" Then
            MsgBox "É obrigatório o preenchimento do campo Somatória de Quantidades.", vbExclamation, Me.Caption
            txtSomatoriaQuantidades.SetFocus
            Exit Sub
        End If
        
    End If
    '''''''''''''''''''''''''''''''''''''
    'Valida campo Somatoria de Depósitos'
    '''''''''''''''''''''''''''''''''''''
    ' Somatória de Datas automático com parametro
    If g_Parametros.chkSoma Then   ' Se Critica Soma
    
        If Trim(txtSomatoriaDepositos.Text) = "" Then
            MsgBox "É obrigatório o preenchimento do campo Somatória de Depósitos.", vbExclamation, Me.Caption
            txtSomatoriaQuantidades.SetFocus
            Exit Sub
        End If
        
    End If
'
    ''''''''''''''''''''''''''''''
    'Valida campo Somatoria 1+2+3'
    ''''''''''''''''''''''''''''''
    ' Somatória de Datas automático com parametro
    If g_Parametros.chkSoma Then   ' Se Critica Soma
    
        If Trim(txtSomatoriaControle.Text) = "" Then
            MsgBox "É obrigatório o preenchimento do campo Somatória de Controle.", vbExclamation, Me.Caption
            txtSomatoriaControle.SetFocus
            Exit Sub
        End If
        
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                      CONSISTÊNCIA NO CALCULO                                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Percorre o Grid somando as datas, quantidades e valores'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim VetDatas(GrdDatas.Rows - 2, 1)
    For i = 1 To GrdDatas.Rows - 1
        
        ''''''''''''''''''''''''''''''''''''''''
        'Pega data de deposito e começa a somar'
        ''''''''''''''''''''''''''''''''''''''''
        sstr = GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO)
        iSomaData = iSomaData + Format(sstr, "ddmmyy")
        ''''''''''''''''''''''''
        'Pega Quantidade e soma'
        ''''''''''''''''''''''''
        sstr = GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE)
        iQuantidade = iQuantidade + sstr
        ''''''''''''''''''''''''
        'Pega Valor do Deposito'
        ''''''''''''''''''''''''
        sstr = GrdDatas.TextMatrix(i, COL_GRD_VALOR)
        If InStr(sstr, ",") Then
            sstr = Mid(sstr, 1, InStr(sstr, ",") - 1) & Mid(sstr, InStr(sstr, ",") + 1)
        ElseIf InStr(sstr, ".") Then
            sstr = Mid(sstr, 1, InStr(sstr, ".") - 1) & Mid(sstr, InStr(sstr, ".") + 1)
        End If
        iValorDeposito = iValorDeposito + sstr
        
        VetDatas(i - 1, 0) = GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO)
        VetDatas(i - 1, 1) = GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE)
    Next i
    
    dSomatoriaGeral = iSomaData + iQuantidade + iValorDeposito
    
    ' Atribui os valores Calculados aos campos de controle se Critica = False
    
    ' Somatória de Datas automático com parametro
    If Not g_Parametros.chkSoma Then   ' Se Critica Soma
        txtSomatoriaDatas.Text = iSomaData
        txtSomatoriaQuantidades.Text = iQuantidade
        txtSomatoriaDepositos.Text = iValorDeposito
        txtSomatoriaControle.Text = dSomatoriaGeral
    End If
    
    ' Os campos de controle serão calculados pelo Sistema se critica = False
    
    If g_Parametros.chkSoma Then   ' Se Critica Soma
        If Val(txtSomatoriaDatas.Text) <> iSomaData Then
            MsgBox "Valores divergentes no campo somatória de datas.", vbExclamation, Me.Caption
            txtSomatoriaDatas.SetFocus
            Exit Sub
        End If
    End If

    If g_Parametros.chkSoma Then   ' Se Critica Soma
        If Val(txtSomatoriaQuantidades.Text) <> iQuantidade Then
            MsgBox "Valores divergentes no campo somatória de quantidades.", vbExclamation, Me.Caption
            txtSomatoriaQuantidades.SetFocus
            Exit Sub
        End If
    End If

    If g_Parametros.chkSoma Then   ' Se Critica Soma
        If Val(txtSomatoriaDepositos.Text) <> iValorDeposito Then
            MsgBox "Valores divergentes no campo somatória dos depósitos.", vbExclamation, Me.Caption
            txtSomatoriaDepositos.SetFocus
            Exit Sub
        End If
    End If

    If g_Parametros.chkSoma Then   ' Se Critica Soma
        If Format(Val(txtSomatoriaControle.Text), "00000000000000") <> Format(dSomatoriaGeral, "00000000000000") Then
            MsgBox "Valores divergentes no campo somatória geral.", vbExclamation, Me.Caption
            txtSomatoriaControle.SetFocus
            Exit Sub
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                INSERÇÃO NA BASE DE DADOS              '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error GoTo BD_Erro_Confirmar
    
    g_cMainConnection.BeginTrans
    
    
    ''''''''''''''''
    'INSERE BORDERO'
    ''''''''''''''''
    If m_Modo = eModo_Inclusao Then
        Call g_cMainConnection.Execute(Proc_Inserir.InsereBordero( _
                                       m_DataProcessamento, _
                                       txtNumeroBordero.Text, _
                                       txtAgencia.Text, _
                                       txtContaCorrente.Text, _
                                       lstTipoCustodia.ItemData(lstTipoCustodia.ListIndex), _
                                       txtLoja.Text, _
                                       Format(txtDataEntradaBordero.MaskText, "yyyymmdd"), _
                                       txtNomeCliente.Text, _
                                       "2", _
                                       txtSomatoriaDatas.Text, _
                                       txtSomatoriaQuantidades.Text, _
                                       txtSomatoriaDepositos.Text, _
                                       txtSomatoriaControle.Text), lRetorno, adCmdText)

        If lRetorno = 0 Then
            g_cMainConnection.RollbackTrans
            MsgBox "Não foi possível inserir o borderô.", vbExclamation, Me.Caption
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Busca o IdBordero para inserir na tabela DataDeposito'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetNumeroBordero(m_DataProcessamento, txtNumeroBordero.Text))
        m_Num_Bordero = txtNumeroBordero.Text
        m_IdBordero = rst!IdBordero
        If rst.EOF() Then
            g_cMainConnection.RollbackTrans
            MsgBox "Não foi possível encontrar o número de Borderô " & _
                    FormataString(txtNumeroBordero.Text, "0", 19, True) & ".", vbExclamation, Me.Caption
            Exit Sub
        End If
    End If
    
    ''''''''''''''''
    'ALTERA BORDERO'
    ''''''''''''''''
    If m_Modo = eModo_Alteracao Then
    
        ''''''''''''''''''''''''''''''''''
        'Seleciona o bordero com o número'
        'que está no campo texto         '
        ''''''''''''''''''''''''''''''''''
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetNumeroBordero( _
                                            m_DataProcessamento, _
                                            txtNumeroBordero.Text))

        ''''''''''''''''''''''''''''''''''''''''''''''
        'Se encontrou um borderô com este número     '
        'e o IdBordero é diferente do IdBordero      '
        'carregado, então não pode continuar, ou seja'
        'não pode alterar o número do bordero para   '
        'um número de borderô já existente           '
        ''''''''''''''''''''''''''''''''''''''''''''''
        If Not rst.EOF() Then
            If m_IdBordero <> rst!IdBordero Then
                Screen.MousePointer = vbDefault
                g_cMainConnection.RollbackTrans
                MsgBox "Não é permitido alterar o número do borderô." & Chr(10) & "Borderô já existente.", vbExclamation, Me.Caption
                txtNumeroBordero.SetFocus
                Exit Sub
            End If
        End If
    
    
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Aqui é utilizado m_IdBordero porque veio da rotina SetIdBordero.'
        'Portanto não precisa ir buscar novamente na base                '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       txtNumeroBordero.Text, _
                                       txtAgencia.Text, _
                                       txtContaCorrente.Text, _
                                       lstTipoCustodia.ItemData(lstTipoCustodia.ListIndex), _
                                       txtLoja.Text, _
                                       Format(txtDataEntradaBordero.MaskText, "yyyymmdd"), _
                                       txtNomeCliente.Text, _
                                       txtSomatoriaDatas.Text, _
                                       txtSomatoriaQuantidades.Text, _
                                       txtSomatoriaDepositos.Text, _
                                       txtSomatoriaControle.Text))

        m_Num_Bordero = txtNumeroBordero.Text
    End If
    
    '''''''''''''''''''''''''''''''
    'Seleciona as datas existentes'
    '''''''''''''''''''''''''''''''
    If m_Modo = eModo_Alteracao Then
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetDatasBordero( _
                                            m_DataProcessamento, _
                                            m_IdBordero))

        For i = 1 To GrdDatas.Rows - 1
        
            rst.Find "DataDeposito = " & Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd")
            
            If (rst.EOF()) Or (m_Consistencia.sDatasExistentes(i - 1).Alteracao) Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Caso não encontrou no rst, então procura no array para ver se houve alteração'
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                For j = 0 To UBound(m_Consistencia.sDatasExistentes)
                    If m_Consistencia.sDatasExistentes(j).Data_Deposito = Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd") And _
                       m_Consistencia.sDatasExistentes(j).Alteracao = True And _
                       m_Consistencia.sDatasExistentes(j).Antiga_Data <> 0 Then
                    
                        ''''''''''''''''''''''''''''''''''''
                        'Faz processo de UPDATE nos Cheques'
                        ''''''''''''''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaDataDepositoCheques( _
                                                       m_DataProcessamento, _
                                                       m_IdBordero, _
                                                       m_Consistencia.sDatasExistentes(j).Antiga_Data, _
                                                       Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd")), _
                                               lRetorno, _
                                               adCmdText)
                        If lRetorno = 0 Then
                            'Não existe cheque para ser atualizado !!!
                            'portanto não é motivo para dar rollback
'                            g_cMainConnection.RollbackTrans
'                            MsgBox "Não foi possível alterar a Data de Deposito dos Cheques.", vbExclamation, Me.Caption
'                            Exit Sub
                        End If
                        '''''''''''''''''''''''''''''''''''''''''
                        'Faz processo de UPDATE nas DataDeposito'
                        '''''''''''''''''''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaDataDeposito( _
                                                       m_DataProcessamento, _
                                                       m_IdBordero, _
                                                       m_Consistencia.sDatasExistentes(j).Antiga_Data, _
                                                       Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd"), _
                                                       GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE), _
                                                       InserePonto(RetiraPonto(GrdDatas.TextMatrix(i, COL_GRD_VALOR)))), _
                                               lRetorno, _
                                               adCmdText)
                        If lRetorno = 0 Then
                            g_cMainConnection.RollbackTrans
                            MsgBox "Não foi possível alterar a Data de Deposito do Borderô.", vbExclamation, Me.Caption
                            Exit Sub
                        End If
                    ''''''''''''''''''''''''''''''''''''''''
                    'Processa inclusão de Datas de Deposito'
                    ''''''''''''''''''''''''''''''''''''''''
                    ElseIf m_Consistencia.sDatasExistentes(j).Data_Deposito = Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd") And _
                       m_Consistencia.sDatasExistentes(j).Inclusao = True Then
                        ''''''''''''''
                        'INSERE DATAS'
                        ''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Inserir.InsereDataDeposito( _
                                                       m_DataProcessamento, _
                                                       m_IdBordero, _
                                                       Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd"), _
                                                       GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE), _
                                                       ColocaPonto(GrdDatas.TextMatrix(i, COL_GRD_VALOR))))
                    End If
                Next j
            End If
            
            If rst.RecordCount > 0 Then rst.MoveFirst
        Next i
        
        On Error Resume Next
        
        i = UBound(m_DatasExclusao)
        
        If Err <> 0 Then i = -1
        
        On Error GoTo Erro_Confirmar:
        
        ''''''''''''''''''''''''''''''''''''''''''''
        'So vai verificar se houver indice no array'
        ''''''''''''''''''''''''''''''''''''''''''''
        If i >= 0 Then
            ''''''''''''''''''''''''''''''''''''''''
            'Varre as datas procurando por exclusão'
            ''''''''''''''''''''''''''''''''''''''''
            For j = 0 To UBound(m_DatasExclusao)
                If m_DatasExclusao(j).Exclusao = True Then
                
                    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetChequesBordero( _
                                                        m_DataProcessamento, _
                                                        m_IdBordero, _
                                                        m_DatasExclusao(j).Data_Deposito), _
                                                lRetorno, _
                                                adCmdText)
                    If Not rst.EOF() Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Excluir todos os cheques que possuem esta Data de Deposito'
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Excluir.RemoveCheques( _
                                                       m_DataProcessamento, _
                                                       m_IdBordero, _
                                                       m_DatasExclusao(j).Data_Deposito), _
                                                lRetorno, _
                                                adCmdText)
                        ''''''''''''''''''''''''''''''''''
                        'Se não excluiu, então volta tudo'
                        ''''''''''''''''''''''''''''''''''
                        If lRetorno = 0 Then
                            g_cMainConnection.RollbackTrans
                            MsgBox "Não foi possível excluir os cheques relacionados à esta data.", vbExclamation, Me.Caption
                            Exit Sub
                        End If
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''
                    'Excluir data selecionada da Tabela DataDeposito'
                    '''''''''''''''''''''''''''''''''''''''''''''''''
                    Call g_cMainConnection.Execute(Proc_Excluir.RemoveDataDeposito( _
                                                   m_DataProcessamento, _
                                                   m_IdBordero, _
                                                   m_DatasExclusao(j).Data_Deposito), _
                                            lRetorno, _
                                            adCmdText)
                    If lRetorno = 0 Then
                        g_cMainConnection.RollbackTrans
                        MsgBox "Não foi possível excluir a Data de Depósito selecionada.", vbExclamation, Me.Caption
                        Exit Sub
                    End If
                End If
            Next j
        End If
    Else
        For i = 1 To GrdDatas.Rows - 1
            ''''''''''''''
            'INSERE DATAS'
            ''''''''''''''
            Call g_cMainConnection.Execute(Proc_Inserir.InsereDataDeposito( _
                                           m_DataProcessamento, _
                                           m_IdBordero, _
                                           Format(GrdDatas.TextMatrix(i, COL_GRD_DATADEPOSITO), "yyyymmdd"), _
                                           GrdDatas.TextMatrix(i, COL_GRD_QUANTIDADE), _
                                           ColocaPonto(GrdDatas.TextMatrix(i, COL_GRD_VALOR))))
        Next i
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Define o IdBordero que acabou de ser incluido'
    '''''''''''''''''''''''''''''''''''''''''''''''
    m_IdBorderoIncluso = Val(m_IdBordero)
    If m_Modo = eModo_Inclusao Then
        m_IdBorderoIncluso = rst!IdBordero
        rst.Close
    End If
    
    
    Set Proc_Inserir = Nothing
    Set Proc_Selecionar = Nothing
    
    ''''''''''''''''''''''
    'Finaliza a Transação'
    ''''''''''''''''''''''
    g_cMainConnection.CommitTrans
    
    ''''''''''''''''''''''''
    'Define retorno da tela'
    ''''''''''''''''''''''''
    
    m_RetornoBordero = eRetornoOK
    
    Unload Me
    
    Exit Sub
    
Erro_Confirmar:
    
    Exit Sub
    
BD_Erro_Confirmar:
    g_cMainConnection.RollbackTrans
    
    TratamentoErro "Não foi possível confirmar o borderô.", Err

End Sub

Private Sub cmdExcluirData_Click()

    ExcluirData

End Sub

Private Sub cmdInserirData_Click()

     Dim eRetornoGrid    As eRetornoGrid
     Dim lret            As Long
     
     If Not ValidaCamposData() Then
         txtDataDeposito.SetFocus
         Exit Sub
     End If
    
     eRetornoGrid = IncluiNoGrid()
    
     If eRetornoGrid = eGR_Erro Then
         MsgBox "Não foi possível inserir na tabela de Datas.", vbExclamation, Me.Caption
     ElseIf eRetornoGrid = eGR_Ok Then
          lret = SendMessage(Me.GrdDatas.hwnd, WM_VSCROLL, SB_BOTTOM, 0)
          LimpaCamposData
          txtDataDeposito.SetFocus
         Exit Sub
     End If

End Sub

Private Sub cmdProvaZero_Click()

' Não está sendo utilizado nesta fase
'    Dim Proc_Alterar        As New Custodia.Atualizar
'    Dim lRetorno            As Long
'
'    ''''''''''''''''''''''''''''''''''''''''''''
'    'Se modo em alteração ja existe m_IdBordero'
'    'caso contrario não existe nenhum IdBordero'
'    ''''''''''''''''''''''''''''''''''''''''''''
'    If m_Modo = eModo_Alteracao Then
'        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
'                                       m_DataProcessamento, _
'                                       m_IdBordero, _
'                                       "4"), _
'                               lRetorno, _
'                               adCmdText)
'        If lRetorno = 0 Then
'            MsgBox "Não foi possível enviar o Borderô para Prova Zero.", vbExclamation, Me.Caption
'            Exit Sub
'        End If
'        LimpaTelaBordero
'    Else
'        MsgBox "Não é possível enviar este Borderô para Prova Zero.", vbExclamation, Me.Caption
'    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

     Dim eRetornoGrid     As eRetornoGrid
     Dim lret            As Long
     
    If KeyAscii = vbKeyReturn Then
    
        Select Case LCase(Screen.ActiveControl.Name)
            Case "txtvalordeposito"
            
                '''''''''''''''''''''''''''''''''''''''''
                'Ao pressionar enter no grid de Datas   '
                'não pode passar para o proximo campo e '
                'sim voltar o foco para data do Deposito'
                '''''''''''''''''''''''''''''''''''''''''
                If Not ValidaCamposData() Then
                    KeyAscii = 0
                    txtDataDeposito.SetFocus
                    Exit Sub
                End If
                
                eRetornoGrid = IncluiNoGrid()
                
                If eRetornoGrid = eGR_Erro Then
                    MsgBox "Não foi possível inserir na tabela de Datas.", vbExclamation, Me.Caption
                ElseIf eRetornoGrid = eGR_Ok Then
                    lret = SendMessage(Me.GrdDatas.hwnd, WM_VSCROLL, SB_BOTTOM, 0)
                    LimpaCamposData
                    txtDataDeposito.SetFocus
                    KeyAscii = 0
                    Exit Sub
                End If
                
                'Caso contrario manda para o proximo controle
                
            Case "txttotal_3"
                '''''''''''''''''''''''''''''''''''''''''''''''''
                'O mesmo quando ocorrer foco no campo txtTotal_3'
                '''''''''''''''''''''''''''''''''''''''''''''''''
                txtSomatoriaDatas.SetFocus
                KeyAscii = 0
                Exit Sub
        End Select
        ''''''''''''''''''''''''''''''''''''''
        'Passa o foco para o proximo controle'
        ''''''''''''''''''''''''''''''''''''''
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Load()

    Dim rst         As New ADODB.Recordset
    Dim Proc        As New Custodia.Selecionar
    
    '''''''''''''''
    'Zera variavel'
    '''''''''''''''
    m_Num_Bordero = ""
    
    
    '''''''''''''''''''''''''''''''''
    'Formata a Data de Processamento'
    '''''''''''''''''''''''''''''''''
    m_DataProcessamento = Geral.DataProcessamento
    ''''''''''''''''''''''''''''
    'Segura as cores de seleção'
    ''''''''''''''''''''''''''''
    m_bColor = GrdDatas.BackColorSel
    m_fColor = GrdDatas.ForeColorSel
    
    '''''''''''''''''''''''
    'Logo retira a seleção'
    '''''''''''''''''''''''
    GrdDatas.BackColorSel = GrdDatas.BackColor
    GrdDatas.ForeColorSel = GrdDatas.ForeColor
    
    '''''''''''''''''''''''''''
    'Determina a conexão ativa'
    '''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc.GetCarteira())
    
    Do While Not rst.EOF()
        lstTipoCustodia.AddItem rst!Descricao
        lstTipoCustodia.ItemData(lstTipoCustodia.NewIndex) = rst!CodigoCarteira
        rst.MoveNext
    Loop
    
    '''''''''''''''''''
    'Fecha o Recordset'
    '''''''''''''''''''
    rst.Close
    Set Proc = Nothing

    Erase m_Consistencia.sDatasExistentes
    Erase m_DatasExclusao
    m_Consistencia.iQuantidade_Cheques_Grid = 0
    m_Consistencia.iQuantidade_Linhas_Grid = 0

    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se é para carregar um borderô já existente'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    m_Modo = eModo_Inclusao
    
    If m_StartarBordero Then
        If Not MostraBordero(m_IdBordero) Then
            MsgBox "Não foi possível iniciar o Borderô.", vbExclamation, Me.Caption
            m_StartarBordero = False
            Exit Sub
        End If
        m_Modo = eModo_Alteracao
        
        'Não permitir a alteração da data de entrada do bordero
        txtDataEntradaBordero.Locked = True
    Else
        'Habilitar o campo "Data de Entrada do Bordero
        txtDataEntradaBordero.Locked = False
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_StartarBordero = False
End Sub

Private Sub GrdDatas_Click()

    If Not (GrdDatas.Rows > 1) Then Exit Sub
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Antes de começar, limpar todas as pendências de alteração'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LimpaPendenciaGrid

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Ao clicar no grid, colocar os dados nos campos para permitir a alteração'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtDataDeposito.Text = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_DATADEPOSITO), "dd/mm/yyyy")
    txtQuantidade.Text = GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_QUANTIDADE)
    txtValorDeposito.Text = GrdDatas.TextMatrix(GrdDatas.Row, COL_GRD_VALOR)
    
    '''''''''''''''''''''''
    'FLAG para atualização'
    '''''''''''''''''''''''
    GrdDatas.RowData(GrdDatas.Row) = 1
    
    ''''''''''''''''''''''''
    'Volta a cor de seleção'
    ''''''''''''''''''''''''
    GrdDatas.BackColorSel = m_bColor
    GrdDatas.ForeColorSel = m_fColor
    
    txtDataDeposito.SetFocus


End Sub
Private Sub GrdDatas_LeaveCell()
Dim A
A = 1
End Sub



Private Sub lstTipoCustodia_Click()

    If m_Event Then Exit Sub
    
    m_Event = True

    SelecionaLista lstTipoCustodia.ListIndex
    
    m_Event = False
End Sub

Private Sub lstTipoCustodia_ItemCheck(Item As Integer)

    If m_Event Then Exit Sub
    
    m_Event = True

    SelecionaLista Item
    
    m_Event = False
End Sub

Private Sub txtDataDeposito_GotFocus()
    SelecionarTexto txtDataDeposito
End Sub
Private Sub txtDataEntradaBordero_LostFocus()

    If Trim(txtDataEntradaBordero.InverseText) > Format(Now, "yyyymmdd") And Len(Trim(txtDataEntradaBordero.InverseText)) <> 0 Then
        MsgBox "Data de Entrada do Borderô não pode ser maior que a data do sistema.", vbExclamation, Me.Caption
        txtDataEntradaBordero.Text = ""
        txtDataEntradaBordero.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtDataEntradaBordero_Validate(Cancel As Boolean)

    'Dim sDataEntrada        As String
    
    
    'sDataEntrada = txtDataEntradaBordero.MaskText
    
    'If Not IsDate(sDataEntrada) Then Exit Sub
    

    'If (Weekday(sDataEntrada, vbSunday) = 1) Or _
    '   (Weekday(sDataEntrada, vbSunday) = 7) Then
       
    '    MsgBox "Não é permitido data de final de semana.", vbInformation
       
    '    Cancel = True

    'End If

End Sub

Private Sub txtLoja_GotFocus()
    SelecionarTexto txtLoja
End Sub
Private Sub txtLoja_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub
Private Sub txtLoja_LostFocus()
    txtLoja.Text = FormataString(txtLoja.Text, "0", txtLoja.MaxLength, True)
End Sub

Private Sub txtNomeCliente_GotFocus()
    SelecionarTexto txtNomeCliente
End Sub
Private Sub txtNumeroBordero_GotFocus()
    SelecionarTexto txtNumeroBordero
End Sub
Private Sub txtNumeroBordero_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub
Private Sub txtNumeroBordero_LostFocus()

    If Trim(txtNumeroBordero.Text) = "" Then Exit Sub

    txtNumeroBordero.Text = FormataString(txtNumeroBordero.Text, "0", txtNumeroBordero.MaxLength, True)

End Sub
Private Sub txtNumeroBordero_Validate(Cancel As Boolean)

    Dim stxtNumeroBordero   As String * 19
    Dim sCPFCGC             As String * 12
    Dim sNumBordero         As String * 6
    Dim sDV                 As String * 1
    
    On Error GoTo NumeroInvalido
    
    If Trim(txtNumeroBordero.Text) = "" Then Exit Sub
    
    stxtNumeroBordero = FormataString(txtNumeroBordero.Text, "0", txtNumeroBordero.MaxLength, True)
    
    sCPFCGC = Left(stxtNumeroBordero, 12)
    
    sNumBordero = Mid(stxtNumeroBordero, 13, 6)
    
    sDV = Right(stxtNumeroBordero, 1)
    
    If Val(sNumBordero) = 0 Then GoTo NumeroInvalido
    
    If Not Modulo11Simplificado(sNumBordero & sDV) Then GoTo NumeroInvalido

    Exit Sub
NumeroInvalido:

    MsgBox "Número de Borderô inválido.", vbExclamation, Me.Caption
    txtNumeroBordero.Text = ""
    Cancel = True
    SelecionarTexto txtNumeroBordero
    Exit Sub

End Sub
Private Sub txtQuantidade_GotFocus()
    SelecionarTexto txtQuantidade
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub
Private Sub txtSomatoriaControle_GotFocus()
    SelecionarTexto txtSomatoriaControle
End Sub
Private Sub txtSomatoriaDatas_GotFocus()
    SelecionarTexto txtSomatoriaDatas
End Sub
Private Sub txtSomatoriaDatas_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub
Private Sub txtSomatoriaDepositos_GotFocus()
    SelecionarTexto txtSomatoriaDepositos
End Sub
Private Sub txtSomatoriaQuantidades_GotFocus()
    SelecionarTexto txtSomatoriaQuantidades
End Sub
Private Sub txtValorDeposito_GotFocus()
    SelecionarTexto txtValorDeposito
End Sub

