VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form DataTroca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data de troca"
   ClientHeight    =   2805
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumRemessa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   468
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1728
      Visible         =   0   'False
      Width           =   2004
   End
   Begin VB.ComboBox cboBanco 
      Height          =   288
      Left            =   3264
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   336
      Width           =   948
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   348
      Left            =   2400
      TabIndex        =   3
      Top             =   2292
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   348
      Left            =   924
      TabIndex        =   2
      Top             =   2292
      Width           =   972
   End
   Begin DATEEDITLib.DateEdit txtDataProcessamento 
      Height          =   396
      Left            =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   336
      Width           =   2004
      _Version        =   65537
      _ExtentX        =   3535
      _ExtentY        =   698
      _StockProps     =   93
      Text            =   "28032002"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "28032002"
      Locked          =   -1  'True
   End
   Begin DATEEDITLib.DateEdit txtDataTroca 
      Height          =   400
      Left            =   1140
      TabIndex        =   1
      Top             =   1056
      Width           =   2004
      _Version        =   65537
      _ExtentX        =   3535
      _ExtentY        =   706
      _StockProps     =   93
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRemessa 
      AutoSize        =   -1  'True
      Caption         =   "Número da Remessa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   1140
      TabIndex        =   7
      Top             =   1488
      Visible         =   0   'False
      Width           =   1764
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data de troca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   1140
      TabIndex        =   5
      Top             =   852
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data de Processamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   1140
      TabIndex        =   4
      Top             =   108
      Width           =   2016
   End
End
Attribute VB_Name = "DataTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_RetornoModal      As enumRetornoModal
Private m_DataTroca         As Long
Private m_TipoArquivo       As enumTipoArquivoCEL
Private m_TipoCheque        As Integer
Private m_TipoGer           As Integer ' 0 - Geração  1- Regeração
Private m_NumRemessa        As Integer
Private m_NovaRemessa       As Boolean


Public Sub SetTipoArquivoCEL(ByVal pTipoArquivoCEL As enumTipoArquivoCEL)
    m_TipoArquivo = pTipoArquivoCEL
End Sub


Public Sub SetTipoGer(ByVal pGerTipo As Integer)
    m_TipoGer = pGerTipo
End Sub
Public Function ShowModal(ByRef pDataTroca As Long, ByRef pTipoCheque As Integer, Optional ByRef pGerTipo As String, Optional ByRef pNumRemessa As Integer, Optional ByRef pNovaRemessa As Boolean) As enumRetornoModal
    
    txtDataProcessamento.Text = FormataData(Geral.DataProcessamento, DD_MM_AAAA)
    
    Me.Show vbModal
    
    pTipoCheque = m_TipoCheque
    
    pNumRemessa = m_NumRemessa
    
    pNovaRemessa = m_NovaRemessa
    
    If m_RetornoModal = eRetornoOK Then
        pDataTroca = m_DataTroca
    End If
    
    ShowModal = m_RetornoModal
    
End Function


Private Sub cmdCancelar_Click()
    m_RetornoModal = eRetornoCancelar
    
    Unload Me
End Sub

Private Sub cmdOK_Click()

    m_RetornoModal = eRetornoOK
    
    
    If Trim(txtDataTroca.Text) = "" Then
        MsgBox "Data Informada Inválida.", vbExclamation, Me.Caption
        txtDataTroca.SetFocus
        Exit Sub
    End If
    
       
    
    If Not IsDate(FormataData(txtDataTroca.InverseText, DD_MM_AAAA)) Then
        MsgBox "Data Informada Inválida.", vbExclamation, Me.Caption
        txtDataTroca.SetFocus
        Exit Sub
    End If
    
    If txtNumRemessa.Visible Then  ' Regeração
    
        If Trim(txtNumRemessa.Text) = "" Or Not IsNumeric(Trim(txtNumRemessa.Text)) Then
            MsgBox "Número da Remessa Inválido.", vbExclamation, Me.Caption
            txtNumRemessa.SetFocus
            Exit Sub
        End If
    
    End If
    
    
    m_TipoCheque = 0
    
    If m_TipoArquivo = eCheque_Unibanco Then
        If Val(cboBanco.ItemData(cboBanco.ListIndex)) = 0 Then
            MsgBox "Selecione um Banco.", vbExclamation, Me.Caption
            cboBanco.SetFocus
            Exit Sub
        End If
        m_TipoCheque = cboBanco.ItemData(cboBanco.ListIndex)
    End If
    
    m_DataTroca = txtDataTroca.InverseText
    
    
    
    If txtNumRemessa.Visible Then
        
       m_NumRemessa = CInt(txtNumRemessa.Text)
       
       If MsgBox("Gerar Novo Número de Remessa?", vbYesNo + vbQuestion) = vbYes Then
            m_NovaRemessa = True
       Else
            m_NovaRemessa = False
       End If
        
    End If
    
    Unload Me
    
End Sub


Private Sub Form_Load()

    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim rst                 As New ADODB.Recordset
    
    txtNumRemessa.Visible = CBool(CStr(m_TipoGer) = CStr(1))
    lblRemessa.Visible = CBool(CStr(m_TipoGer) = CStr(1))
    cboBanco.Visible = CBool(m_TipoArquivo = eCheque_Unibanco)
    
    If m_TipoGer = 1 Then
        DataTroca.Caption = "Regeração de Remessa"
    Else
        DataTroca.Caption = "Geração de Remessa"
    End If
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Só preenche o combo se tipo de arquivo fora definido como unibanco'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If m_TipoArquivo = eCheque_Unibanco Then
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetTipoCheque())
        Do While Not rst.EOF
            '''''''''''''''''''''''''''''''''''''
            'Só coloca tipocheque diferente de 0'
            '''''''''''''''''''''''''''''''''''''
            If rst!TipoCheque <> 0 Then
                cboBanco.AddItem rst!TipoCheque & " - " & rst!Descricao
                cboBanco.ItemData(cboBanco.NewIndex) = rst!TipoCheque
            End If
            rst.MoveNext
        Loop
        rst.Close
        '''''''''''''''''''''''''''''''
        'Seleciona o primeiro da lista'
        '''''''''''''''''''''''''''''''
        If cboBanco.ListCount > 0 Then cboBanco.ListIndex = 0
        
    ElseIf m_TipoArquivo = eArquivo_TER Then
        
    End If
    
    Set Proc_Selecionar = Nothing
        

End Sub

Private Sub txtDataTroca_KeyPress(KeyAscii As Integer)
    If txtNumRemessa.Visible Then
        If KeyAscii = vbKeyReturn Then
            txtNumRemessa.SetFocus
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            cmdOK_Click
        End If
    End If
End Sub

Private Sub txtNumRemessa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       cmdOK_Click
    End If
End Sub
