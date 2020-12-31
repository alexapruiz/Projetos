VERSION 5.00
Begin VB.Form FrmDelEnvMal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão de Envelope / Malote"
   ClientHeight    =   1584
   ClientLeft      =   288
   ClientTop       =   1572
   ClientWidth     =   11460
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1584
   ScaleWidth      =   11460
   Begin VB.Frame Frame3 
      Height          =   1260
      Left            =   0
      TabIndex        =   7
      Top             =   108
      Width           =   9456
      Begin VB.TextBox TxtNumMalote 
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
         TabIndex        =   1
         Top             =   228
         Width           =   2196
      End
      Begin VB.ComboBox CmbAgencia 
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
         ItemData        =   "FrmDelEnvMal.frx":0000
         Left            =   2280
         List            =   "FrmDelEnvMal.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   732
         Width           =   2604
      End
      Begin VB.PictureBox Picture6 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   732
         Width           =   2100
         Begin VB.Label Label5 
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
            TabIndex        =   13
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
         TabIndex        =   0
         Top             =   240
         Width           =   2604
      End
      Begin VB.PictureBox Picture4 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label Label4 
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
            TabIndex        =   11
            Top             =   12
            Width           =   1992
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   396
         Left            =   4944
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label Label3 
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
            TabIndex        =   9
            Top             =   36
            Width           =   1956
         End
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   396
         Left            =   4944
         TabIndex        =   15
         Top             =   732
         Width           =   2100
      End
      Begin VB.Label lblLote 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   7116
         TabIndex        =   14
         Top             =   732
         Width           =   2196
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   9540
      TabIndex        =   6
      Top             =   108
      Width           =   1752
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   156
         TabIndex        =   5
         Top             =   876
         Width           =   1464
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   324
         Left            =   144
         TabIndex        =   3
         Top             =   180
         Width           =   1464
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "E&xclusão"
         Height          =   324
         Left            =   144
         TabIndex        =   4
         Top             =   528
         Width           =   1464
      End
   End
End
Attribute VB_Name = "FrmDelEnvMal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsAux        As rdoResultset
Dim RsMot        As rdoResultset               'Recordset Motivo de Exclusão
Dim RsStatus     As rdoResultset
Dim RsPesqOcorr  As rdoResultset
Dim FlagExclusao As String                     'Flag de Ajuste de Tela

Dim TEnvMal     As String

Private Type MdregOcorrencia
    qryGetLstOcorrencias    As rdoQuery
    qryGetPesqMalote        As rdoQuery
    qryGetVerStatus         As rdoQuery
End Type
Private MdregOcorrencia  As MdregOcorrencia


Private Type MdSupExclusao
    qryGetPesqCapa        As rdoQuery
    qryGetAgCapa As rdoQuery
    qryGetPesqMotivExclus As rdoQuery
    qryGetMostraStatus    As rdoQuery
End Type
Private MdSupExclusao As MdSupExclusao
Private Sub limpa_Header()
    lblLote.Caption = ""
    cmbCapa.Clear
    TxtNumMalote.Text = ""
    CmbAgencia.Clear
End Sub
Private Sub cmbAgencia_Change()

    If CmbAgencia.ListCount = 1 Then
        CmbAgencia.Text = CmbAgencia.List(0)
        Call CmdExcluir_Click
    End If
    
End Sub
Private Sub cmbAgencia_Click()
'Muda o Número de Lote - Quando for Selecionado outra Agência

    If CmbAgencia.Text = "" Then Exit Sub

    With MdSupExclusao.qryGetAgCapa
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = CmbAgencia
        .rdoParameters(3).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAux.EOF Then
        lblLote = Format(RsAux!IdLote, "0000-00000")
        Call CmdExcluir_Click
    End If
  
End Sub
Private Sub cmdLimpaCampos_Click()
    LimpaTela Me
    txtNumEnvMal.SetFocus
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub OptCapaEnvelope_Click()
    'Posiciona o foco no Text - txtnumEnvmal
    txtNumEnvMal.SetFocus
End Sub
Private Sub OptCapaMalote_Click()
    'Posiciona o foco no Text - txtnumEnvmal
    txtNumEnvMal.SetFocus
End Sub
Private Sub TxtNumEnvMal_KeyPress(KeyAscii As Integer)

 InibirTeclaAlfa KeyAscii

 If (KeyAscii = 13) Then
        
        If Len(txtNumEnvMal) > 0 Then
            If OptCapaEnvelope.Value Then
            txtNumEnvMal = Format(txtNumEnvMal, "00000000")
            Else
            txtNumEnvMal = Format(txtNumEnvMal, "00000000000000")
            End If
       End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click
    End If
    
End Sub
Private Sub TxtNumEnvMal_LostFocus()
'* Formatação de Número de Envelope e Malote *'

    If OptCapaEnvelope.Value Then
        txtNumEnvMal = Format(txtNumEnvMal, "00000000")
    Else
        txtNumEnvMal = Format(txtNumEnvMal, "00000000000000")
    End If
                        
End Sub
Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If CmbAgencia.ListCount >= 1 Then
            'cmbAgencia.Text = cmbAgencia.List(cmbAgencia.ListIndex)
            Call CmdExcluir_Click
        End If
    End If
    
End Sub
Private Sub cmbCapa_Change()
    If Len(TxtNumMalote) <> 0 Then
        If Len(cmbCapa) <> 0 And cmbCapa.ListCount > 1 Then
            Call Pesquisa_Dados
        End If
    End If
End Sub
Private Sub cmbCapa_GotFocus()
    SelecionarTexto cmbCapa
End Sub

Private Sub cmbCapa_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Then
      If Len(cmbCapa) > 0 Then
         Call Pesquisa_Dados
      End If
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   Else
      'Não permitir a digitação de mais de 18 caracteres
      If Len(cmbCapa.Text) >= 18 And cmbCapa.SelLength = 0 And (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdLimpar_Click()
    
    CmbAgencia.Clear
    cmbCapa.Clear
    lblLote = ""
    LimpaTela Me
    If cmbCapa.Enabled Then cmbCapa.SetFocus

End Sub
Private Sub CmdExcluir_Click()
   
Dim TotalDoctos, TotalDoctoComOCorr, TotalDoctosSemCorr, TotalGeral As Integer
Dim lIdCapa As Long

    If FlagExclusao = "OK" Then
       FlagExclusao = ""
       Exit Sub
    End If
    
    'Valida Capa
    If Trim(cmbCapa.Text) = "" Then
        MsgBox "Campo Obrigatório não preenchido!", vbInformation, App.Title
        cmbCapa.SetFocus
        Exit Sub
    End If

    'Valida Agência
    If Trim(CmbAgencia.Text) = "" Then
        MsgBox "Campo Obrigatório não preenchido!", vbInformation, App.Title
        CmbAgencia.SetFocus
        Exit Sub
    End If

    With MdSupExclusao.qryGetPesqCapa
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
        .rdoParameters(2).Value = TEnvMal
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsAux.EOF Then
        'Recupera IdCapa da Capa Selecionada
        lIdCapa = RsAux!IdCapa

           'Verificar se a capa pode ser excluida
           If Not VerificaDoctosExcluidosCapa(lIdCapa) Then
               MsgBox "Não é permitido excluir Envelopes / Malotes em que todos os documentos possuam ocorrência.", vbInformation + vbOKOnly, App.Title
               Exit Sub
           End If

            With MdSupExclusao.qryGetPesqMotivExclus
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = lIdCapa
                Set RsMot = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With

            With MdSupExclusao.qryGetMostraStatus
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = Null
                .rdoParameters(2).Value = Null
                .rdoParameters(3).Value = CDbl(cmbCapa)
                
                Set RsStatus = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With

            If RsAux!Status = "D" Then
                MotivoExclusao.LblValorEnv_Mal = cmbCapa.Text
                MotivoExclusao.LblValorEnv_Mal.Tag = lIdCapa
                MotivoExclusao.TxtMotivo = RsMot!MotivoExclusao
                MotivoExclusao.CmdOK.Enabled = True
                MotivoExclusao.Show vbModal, Me

            ElseIf RsAux!Status = "F" Then
                MotivoExclusao.LblValorEnv_Mal = cmbCapa.Text
                MotivoExclusao.LblValorEnv_Mal.Tag = lIdCapa
                MotivoExclusao.TxtMotivo = RsMot!MotivoExclusao
                MotivoExclusao.CmdOK.Enabled = False
                MotivoExclusao.Show vbModal, Me
                Exit Sub

            ElseIf RsAux!Status = "P" Or RsAux!Status = "X" Then
                MsgBox "Registro Excluido - " & RsStatus!Descricao, vbInformation, App.Title
                Exit Sub

            ElseIf RsAux!Status = "2" Or RsAux!Status = "3" Or RsAux!Status = "E" _
                Or RsAux!Status = "G" Or RsAux!Status = "H" Or RsAux!Status = "I" _
                Or RsAux!Status = "J" Or RsAux!Status = "K" Or RsAux!Status = "S" _
                Or RsAux!Status = "T" Or RsAux!Status = "N" Or RsAux!Status = "Q" Then

                If TEnvMal = "M" Then
                    MsgBox "Este Malote não está disponível. Pode estar sendo tratado por outra estação ou já foi tratado.", vbInformation, App.Title
                Else
                    MsgBox "Este Envelope não está disponível. Pode estar sendo tratado por outra estação ou já foi tratado.", vbInformation, App.Title
                End If
                Exit Sub
            
            Else
                MotivoExclusao.LblValorEnv_Mal = cmbCapa.Text
                MotivoExclusao.LblValorEnv_Mal.Tag = lIdCapa
                MotivoExclusao.TxtMotivo = ""
                MotivoExclusao.Show vbModal, Me
            End If

            'Verificar Retorno
            If MotivoExclusao.Result = True Then
            'Excluiu -> Gravar Log
              Call GravaLog(lIdCapa, 0, 100)
              cmdLimpar_Click
              FlagExclusao = "OK"
            Else
              cmdLimpar_Click
              FlagExclusao = "OK"
            End If
        End If
        
End Sub
Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(21)
   Call cmdLimpar_Click
End Sub
Private Sub Form_Load()

'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
Set MdSupExclusao.qryGetPesqCapa = Geral.Banco.CreateQuery("", "{Call GetTodasCapas(?,?,?)}")

'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
Set MdregOcorrencia.qryGetPesqMalote = Geral.Banco.CreateQuery("", "{Call GetMaloteExpedicao(?,?)}")

'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
Set MdSupExclusao.qryGetAgCapa = Geral.Banco.CreateQuery("", "{Call GetAgenciasCapa(?,?,?,?)}")

'Traz o Motivo de Exclusão para o Registro Corrente
Set MdSupExclusao.qryGetPesqMotivExclus = Geral.Banco.CreateQuery("", "{Call GetMotivoExclusao(?,?)}")

'Traz o Motivo de Exclusão para o Registro Corrente
Set MdSupExclusao.qryGetMostraStatus = Geral.Banco.CreateQuery("", "{Call GetrecuperaStatus (?,?,?,?)}")

End Sub

Public Sub Pesquisa_Dados()

Dim CountAg As Integer
Dim CounRegExcluido As Integer
       
    CmbAgencia.Clear
    
    If cmbCapa.Text = "" Then Exit Sub

    'Se Capa Não For diferente de Zero verifica seu dados
    With MdSupExclusao.qryGetAgCapa
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = Null
        .rdoParameters(3).Value = Null
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsAux.EOF Then
       TEnvMal = RsAux!IdEnv_Mal
       For CountAg = 0 To RsAux.RowCount - 1
           CmbAgencia.AddItem RsAux!AgOrig
           CmbAgencia.ItemData(CmbAgencia.NewIndex) = RsAux!IdCapa
    
           If RsAux!Num_Malote <> 0 Then
                Me.TxtNumMalote.Text = FormataMalote(RsAux!Num_Malote)
            End If
           RsAux.MoveNext
       Next
           
        'Abre Combo Com a Lista das Agencias de ListCoun For > que 1
        If CmbAgencia.ListCount = 1 Then
           CmbAgencia.Text = CmbAgencia.List(0)
           Call CmdExcluir_Click
        ElseIf CmbAgencia.ListCount > 1 Then
           CmbAgencia.Text = CmbAgencia.List(0)
           CmbAgencia.SetFocus
           SendKeys "{F4}"
        End If
    
    Else
        MsgBox "Registro não encontrado!", vbInformation, App.Title
        Exit Sub
    End If

End Sub
Public Sub Pesquisa_Malote()

On Error Resume Next

   Dim CountCapaMalote As Integer
   Dim CountRegExcluido As Integer

   TxtNumMalote = FormataMalote(TxtNumMalote)
    
   If Len(TxtNumMalote.Text) = 9 Or Len(TxtNumMalote.Text) = 10 Or Len(TxtNumMalote.Text) = 11 Or Len(TxtNumMalote.Text) = 12 Then
      CmbAgencia.Clear
      cmbCapa.Clear
      RsAux.Close

      With MdregOcorrencia.qryGetPesqMalote
         .rdoParameters(0).Value = Geral.DataProcessamento
         .rdoParameters(1).Value = Val(TxtNumMalote)
        
         If Err = 13 Then
            MsgBox "Valor Invalido, Reentre!", vbInformation, App.Title
            'TxtNumMalote.Text = ""
            TxtNumMalote.SelStart = 0
            TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
            TxtNumMalote.SetFocus
            Exit Sub
         End If
        
         Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
      End With

      If Not RsAux.EOF Then
       
         If RsAux.RowCount = 1 Then
            If RsAux!Status = "P" Or RsAux!Status = "E" Then
               MsgBox "Registro já foi Excluido.", vbInformation, App.Title
               Exit Sub
            End If
         End If

         For CountAg = 0 To RsAux.RowCount - 1
            If RsAux!Status = "P" Or RsAux!Status = "E" Then
               CountRegExcluido = CountRegExcluido + 1
               Call RetiraDuplicidade(RsAux!Capa)
            Else
               cmbCapa.AddItem RsAux!Capa
               Call RetiraDuplicidade(RsAux!Capa)
            End If

            RsAux.MoveNext
            DoEvents
         Next

         If CountRegExcluido = RsAux.RowCount Then
            MsgBox "Registro já foi Excluido.", vbInformation, App.Title
         End If

         If cmbCapa.ListCount = 1 Then
            cmbCapa.Text = cmbCapa.List(0)
         ElseIf cmbCapa.ListCount > 1 Then
            cmbCapa.SetFocus
            SendKeys "{F4}"
         End If

         Call Pesquisa_Dados
      Else
         MsgBox "Registro não encontrado!", vbInformation, App.Title
         TxtNumMalote.SelStart = 0
         TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
         TxtNumMalote.SetFocus
         Exit Sub
      End If
   Else
      MsgBox "Dados inválido, reentre!", vbInformation, App.Title
      TxtNumMalote.SetFocus
      Exit Sub
   End If
   
End Sub
Private Sub TxtNumMalote_GotFocus()
    SelecionarTexto TxtNumMalote
End Sub
Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)

   If (KeyAscii = vbKeyReturn) Then
      If Len(TxtNumMalote.Text) > 0 Then
          If VerificaMalote(TxtNumMalote) = False Then
              MsgBox "Número de Malote inválido.", vbInformation, App.Title
              TxtNumMalote.SelStart = 0
              TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
              TxtNumMalote.SetFocus
              Exit Sub
           End If
           Call Pesquisa_Malote
      End If
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
   
End Sub
Function VerificaMalote(ValNumMalote As String) As Boolean
'* Verifica se número de malote é Válido *'

    If Len(ValNumMalote) = 12 And CStr(Mid(ValNumMalote, 1, 2)) <> "09" Then
       VerificaMalote = False
    Else
       VerificaMalote = True
    End If
        
End Function
Function RetiraDuplicidade(NumCapaMalote As Double)
'* Elimina da Lista de Capas sua duplicidades *'

Dim CountLoop   As Integer  'Conta o Loop de acordo com a quantidade de Capas no Combo
Dim CountCapa   As Integer  'Traz  a quantidade de registros duplidados
Dim GuardaItem  As Integer

    For CountLoop = 0 To cmbCapa.ListCount - 1
        If NumCapaMalote = cmbCapa.List(CountLoop) Then
            CountCapa = CountCapa + 1
            GuardaItem = CountLoop
        End If
        
        If CountCapa >= 2 Then
            cmbCapa.RemoveItem (CountLoop)
        End If
    Next
      
End Function
