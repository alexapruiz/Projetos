VERSION 5.00
Begin VB.Form FinalizacaoCapa 
   Caption         =   "Finalização de Capa"
   ClientHeight    =   1284
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   11568
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1284
   ScaleWidth      =   11568
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   9768
      TabIndex        =   10
      Top             =   -48
      Width           =   1752
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   132
         TabIndex        =   12
         Top             =   528
         Width           =   1464
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   324
         Left            =   132
         TabIndex        =   11
         Top             =   168
         Width           =   1464
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   48
      TabIndex        =   0
      Top             =   -48
      Width           =   9672
      Begin VB.TextBox txtNumMalote 
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
         Height          =   420
         Left            =   7116
         MaxLength       =   12
         TabIndex        =   9
         Top             =   228
         Width           =   2196
      End
      Begin VB.ComboBox cmbAgencia 
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
         Height          =   396
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   732
         Width           =   2604
      End
      Begin VB.PictureBox Picture5 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   732
         Width           =   2100
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Agência"
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
            Height          =   300
            Left            =   -48
            TabIndex        =   7
            Top             =   12
            Width           =   984
         End
      End
      Begin VB.ComboBox cmbCapa 
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
         Height          =   396
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2604
      End
      Begin VB.PictureBox Picture3 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label lblCapa 
            AutoSize        =   -1  'True
            Caption         =   "Capa"
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
            Height          =   312
            Left            =   12
            TabIndex        =   4
            Top             =   12
            Width           =   1992
         End
      End
      Begin VB.PictureBox picNumMalote 
         Height          =   396
         Left            =   4944
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label lblMalote 
            Caption         =   "Número do Malote"
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
            Height          =   252
            Left            =   36
            TabIndex        =   2
            Top             =   36
            Width           =   1956
         End
      End
   End
End
Attribute VB_Name = "FinalizacaoCapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsExpedicao                     As rdoResultset
Private m_IdCapa                        As Long
Private m_Busy                          As Boolean
Private aIndice()                       As Integer
Private m_bEvent                        As Boolean
Private Function FinalizaCapa(ByVal pIdCapa As Double) As Boolean

    Dim qryFinalizaCapa         As rdoQuery
    
    On Error GoTo Erro_FinalizaCapa
    
    FinalizaCapa = False
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Caso exista somente esta capa, entao finaliza'
    '''''''''''''''''''''''''''''''''''''''''''''''
    Set qryFinalizaCapa = Geral.Banco.CreateQuery("", "{? = Call AtualizaCapaFinalizada(?,?,?)}")
    With qryFinalizaCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = pIdCapa
        .rdoParameters(3) = 1
        .Execute
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível finalizar a capa " & cmbCapa.Text, vbExclamation
            'rsLocalizaCapa.Close
            Exit Function
        End If
    End With
    
    FinalizaCapa = True
    
    Exit Function
    
Erro_FinalizaCapa:

    Select Case TratamentoErro("Erro ao finalizar a capa.", Err, rdoErrors)
        Case vbRetry
            Resume
        Case vbCancel
    End Select
    

End Function

Private Sub LimpaTela()

    cmbCapa.Text = ""
    cmbAgencia.Clear
    txtNumMalote.Text = ""
    
End Sub

Private Function LocalizaCapa(ByVal pCapa As Double, ByVal pIdEnv_Mal As String) As Boolean

    Dim rsLocalizaCapa          As rdoResultset
    Dim qryGetFinalizados       As rdoQuery
    
    On Error GoTo Erro_LocalizaCapa
    
    
    LocalizaCapa = False
    
    
    If pIdEnv_Mal = "E" Then
        '''''''''''''''''''''''''''
        'Abre o ResultSet de capas'
        '''''''''''''''''''''''''''
        Set qryGetFinalizados = Geral.Banco.CreateQuery("", "{Call GetCapaFinalizada(?,?,?,?)}")
        With qryGetFinalizados
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = pCapa
            .rdoParameters(2) = IIf(Trim(cmbAgencia.Text) = "", Null, Val(cmbAgencia.Text))
            .rdoParameters(3) = 0
            Set rsLocalizaCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    Else
        '''''''''''''''''''''''''''''
        'Abre o ResultSet de Malotes'
        '''''''''''''''''''''''''''''
        
        If Len(txtNumMalote.Text) = 12 Then
            If Left(txtNumMalote.Text, 2) <> "09" Then
                MsgBox "Número do Malote inválido.", vbExclamation + vbOKOnly, App.Title
                LimpaTela
                Exit Function
            End If
        End If
        
        Set qryGetFinalizados = Geral.Banco.CreateQuery("", "{Call GetMaloteFinalizada(?,?,?,?)}")
        With qryGetFinalizados
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = pCapa
            .rdoParameters(2) = IIf(Trim(cmbAgencia.Text) = "", Null, Val(cmbAgencia.Text))
            .rdoParameters(3) = 0
            Set rsLocalizaCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    End If
    If rsLocalizaCapa.EOF Then
        MsgBox "Não foi possível encontrar a capa solicitada.", vbExclamation
        rsLocalizaCapa.Close
        Exit Function
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''
    'Preenche o combo agencia com a agencia'
    ''''''''''''''''''''''''''''''''''''''''
    cmbAgencia.Clear
    Do While Not rsLocalizaCapa.EOF
    
        cmbAgencia.AddItem rsLocalizaCapa!AgOrig
        cmbAgencia.ItemData(cmbAgencia.NewIndex) = rsLocalizaCapa!IdCapa
        
        rsLocalizaCapa.MoveNext
    Loop

    ''''''''''''''''''''''''''''''''''''''''''''''
    'Existe mais de uma agencia para a mesma capa'
    ''''''''''''''''''''''''''''''''''''''''''''''
    LocalizaCapa = True
    
    If cmbAgencia.ListCount > 1 Then
        cmbAgencia.SetFocus
        SendKeys "{F4}"
        LocalizaCapa = False
    Else
        cmbAgencia.ListIndex = 0
    End If

    rsLocalizaCapa.Close
    
    Exit Function
    
Erro_LocalizaCapa:

    Select Case TratamentoErro("Erro ao localizar esta capa.", Err, rdoErrors)
    Case vbRetry
        Resume
    Case vbCancel
        LimpaTela
    End Select

End Function

Private Sub cmbAgencia_Click()

    Dim sIdEnv_Mal      As String

    If m_bEvent = True Then Exit Sub

    m_bEvent = True
    
    sIdEnv_Mal = "E"
    sIdEnv_Mal = IIf(Trim(txtNumMalote.Text) <> "", "M", "E")

    If LocalizaCapa(Val(cmbCapa.Text), sIdEnv_Mal) Then
        If FinalizaCapa(CDbl(cmbAgencia.ItemData(cmbAgencia.ListIndex))) Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Conseguiu finalizar, entao limpa tela para a proxima'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            LimpaTela
        End If
    End If
    
    m_bEvent = False

End Sub


Private Sub cmbCapa_Change()

    If Not IsNumeric(cmbCapa.Text) Then cmbCapa.Text = ""

End Sub
Private Sub cmbCapa_KeyPress(KeyAscii As Integer)

    If m_bEvent = True Then Exit Sub
    
    m_bEvent = True

    If KeyAscii = vbKeyReturn Then
        If LocalizaCapa(Val(Trim(cmbCapa.Text)), "E") Then
            If FinalizaCapa(CDbl(cmbAgencia.ItemData(cmbAgencia.ListIndex))) Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Conseguiu finalizar, entao limpa tela para a proxima'
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                LimpaTela
            End If
        Else
            LimpaTela
        End If
    End If
    
    m_bEvent = False

    ''''''''''''''''''''''''''''''''''
    'Permite somente numeros no campo'
    ''''''''''''''''''''''''''''''''''
    SoNumero KeyAscii
    
   'Limita Digitacao a 14 digitos
    If Len(cmbCapa.Text) > 13 Then
        KeyAscii = 0
    End If
End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub


Private Sub cmdLimpar_Click()

    lblMalote.Enabled = True
    txtNumMalote.Enabled = True
    
    cmbCapa.Text = ""
    txtNumMalote.Text = ""
    cmbAgencia.Clear
    
    
    cmbCapa.SetFocus
End Sub


Private Sub Form_Activate()

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(26)
    
End Sub

Private Sub txtNumMalote_Change()

    If Not IsNumeric(txtNumMalote.Text) Then txtNumMalote.Text = ""

End Sub


Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)

    If m_bEvent = True Then Exit Sub
    
    m_bEvent = True

    If KeyAscii = vbKeyReturn Then
        If LocalizaCapa(Val(Trim(txtNumMalote.Text)), "M") Then
            If FinalizaCapa(CDbl(cmbAgencia.ItemData(cmbAgencia.ListIndex))) Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Conseguiu finalizar, entao limpa tela para a proxima'
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                LimpaTela
            End If
        Else
            SelecionarTexto txtNumMalote
        End If
    End If
    
    m_bEvent = False

    ''''''''''''''''''''''''''''''''''
    'Permite somente numeros no campo'
    ''''''''''''''''''''''''''''''''''
    SoNumero KeyAscii
    

End Sub


