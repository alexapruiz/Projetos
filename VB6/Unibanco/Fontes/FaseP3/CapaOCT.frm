VERSION 5.00
Begin VB.Form CapaOCT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capa de OCT"
   ClientHeight    =   2148
   ClientLeft      =   360
   ClientTop       =   1524
   ClientWidth     =   9804
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2148
   ScaleWidth      =   9804
   Begin VB.Frame fraCMC7 
      Caption         =   "Linha do CMC7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   3504
      TabIndex        =   10
      Top             =   1152
      Width           =   6012
      Begin VB.TextBox txtCMC72 
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
         Height          =   408
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1236
      End
      Begin VB.TextBox txtCMC73 
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
         Height          =   408
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   2
         Top             =   240
         Width           =   1452
      End
      Begin VB.TextBox txtCMC71 
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
         Height          =   408
         Left            =   816
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1044
      End
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   800
      Left            =   7824
      Picture         =   "CapaOCT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   800
      Left            =   8688
      Picture         =   "CapaOCT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   800
      Left            =   4368
      Picture         =   "CapaOCT.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   800
      Left            =   3504
      Picture         =   "CapaOCT.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5232
      Picture         =   "CapaOCT.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter Cor"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   6096
      Picture         =   "CapaOCT.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Frente/Verso"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   6960
      Picture         =   "CapaOCT.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   96
      Width           =   852
   End
   Begin VB.Label lblInformativo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação da Capa de OCT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   564
      TabIndex        =   11
      Top             =   348
      Width           =   2328
   End
   Begin VB.Image imgInformativo 
      Height          =   384
      Left            =   96
      Picture         =   "CapaOCT.frx":1546
      Top             =   240
      Width           =   384
   End
End
Attribute VB_Name = "CapaOCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qryAtualizaDocumentoExcluido As rdoQuery        'Atualiza Status = "D", Duplicidade = 1, Ocorrencia = 998
Private qryRemoveTipoDocumento As rdoQuery
Private qryAtualizaCapaOCT  As rdoQuery

'Variavel de retorno informando se Cancelou ou Alterou
Public Alterou As Boolean

Private mForm As Form

Private Sub cmdConfirmar_Click()

Dim bDuplicidade As Boolean
Dim sPosicaoErro As String
sPosicaoErro = "InsCapaOCT"

On Error GoTo Err_cmdConfirmar

If ValidarDados Then
    
    Geral.Documento.Leitura = (txtCMC71.Text + txtCMC72.Text + txtCMC73.Text)
    
    'verifica preenchimento de CMC7
    If VerificaPreenchimentoCMC7(Me) = False Then Exit Sub
    
    'Inicia Transação
    Geral.Banco.BeginTrans
    
    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 39 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If
    
    If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , etpdocCapaOCT, Geral.Documento.Leitura, "1") Then
        Beep
        Alterou = False
        Geral.Banco.RollbackTrans
        MsgBox "Não foi possível complementar este documento!", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
' Se Documento For Capa de OCT o Valor deve ser sempre 0
      With qryAtualizaCapaOCT
         .rdoParameters(0).Direction = rdParamReturnValue
         .rdoParameters(1) = Geral.Documento.IdDocto
         .rdoParameters(2) = Geral.DataProcessamento
         .Execute
      End With
    
    If bDuplicidade Then
        Geral.Documento.Duplicidade = 1
        If Not AtualizaDocumentoExcluido(Geral.Documento.IdDocto) Then
            GoTo Err_cmdConfirmar
        End If
    Else
        Geral.Documento.Duplicidade = 0
    End If
    
    'Finaliza Transação
    Geral.Banco.CommitTrans
    
    Alterou = True
    Me.Hide

End If

Exit Sub

Exit_cmdConfirmar:
    Alterou = False
    Geral.Banco.RollbackTrans
    Exit Sub

Err_cmdConfirmar:
    Alterou = False
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível atualizar/inserir o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    Me.Hide

End Sub

Private Sub cmdFrenteVerso_Click()
    
    mForm.cmdFrenteVerso_Click

End Sub

Private Sub cmdInverteCor_Click()
    
    mForm.cmdInverteCor_Click
    
End Sub

Private Sub cmdRotacao_Click()
    
    mForm.cmdRotacao_Click
    
End Sub

Private Sub CmdSair_Click()
    
    Alterou = False
    Me.Hide

End Sub

Private Sub cmdZoomMais_Click()
    
    mForm.cmdZoomMais_Click
    
End Sub


Private Sub cmdZoomMenos_Click()
    
    mForm.cmdZoomMenos_Click
    
End Sub


Private Sub Form_Activate()

    'Evita efeito sendkeys do evento change
    txtCMC71.Tag = False
    txtCMC72.Tag = False
    txtCMC73.Tag = False

    If Len(Geral.Documento.Leitura) = 30 And _
        Mid(Geral.Documento.Leitura, 1, 3) = "409" And _
        Mid(Geral.Documento.Leitura, 20, 3) = "592" And _
        Mid(Geral.Documento.Leitura, 9, 3) = "592" And _
        Mid(Geral.Documento.Leitura, 12, 6) = Mid(Geral.Documento.Leitura, 24, 6) Then
        
        'Formata o Número da Capa de OCT
        txtCMC71.Text = Left(Geral.Documento.Leitura, 8)
        txtCMC72.Text = Mid(Geral.Documento.Leitura, 9, 10)
        txtCMC73.Text = Mid(Geral.Documento.Leitura, 19)
        
    End If

    txtCMC71.SelStart = 0: txtCMC71.SelLength = txtCMC71.MaxLength
    txtCMC71.SetFocus
        
    'Habilita efeito sendkeys do evento change
    txtCMC71.Tag = True
    txtCMC72.Tag = True
    txtCMC73.Tag = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = &H1B Then
        If InStr(Format(ActiveControl.Name, ">"), "TXTCMC7") <> 0 Then Exit Sub
        CmdSair_Click
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      cmdZoomMais_Click
    Case vbKeySubtract
      cmdZoomMenos_Click
    Case vbKeyF10
      cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      cmdRotacao_Click
    Case vbKeyF11
      cmdFrenteVerso_Click
    Case vbKeyMultiply
      Call cmdConfirmar_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
      mForm.Form_KeyUp KeyCode, Shift
  End Select
End Sub
Private Sub Form_Load()
    
    Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoExcluido (?,?,?,?,?)}")
    Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
    Set qryAtualizaCapaOCT = Geral.Banco.CreateQuery("", "{? = call AtualizaCapaOCT (?,?)}")
    
    Alterou = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  qryAtualizaDocumentoExcluido.Close
  qryRemoveTipoDocumento.Close
  
End Sub

Private Sub txtCMC71_Change()
    
    If txtCMC71.Tag Then
        If Len(Trim(txtCMC71.Text)) = txtCMC71.MaxLength Then SendKeys "{ENTER}"
    End If

End Sub

Private Sub txtCMC71_GotFocus()
    
    txtCMC71.SelStart = 0
    txtCMC71.SelLength = txtCMC71.MaxLength

End Sub

Private Sub txtCMC71_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then KeyCode = 0

End Sub

Private Sub txtCMC71_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC71) < 8 Then
            Beep
            MsgBox "Digite todos os números do primeiro campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            txtCMC71.SelStart = 0
            txtCMC71.SelLength = txtCMC71.MaxLength
            GoTo Sair
        End If
        'Verifica o código do documento
        If Left(txtCMC71.Text, 3) <> "409" Then
            Beep
            MsgBox "Para Capa de OCT o primeiro campo do CMC7 deve começar com o Nr. 409.", vbInformation + vbOKOnly, App.Title
            GoTo Sair
        End If
        txtCMC72.SetFocus
        
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
    End If
    
    Exit Sub
    
Sair:
    txtCMC71_GotFocus
    txtCMC71.SetFocus

End Sub

Private Sub txtCMC72_Change()
    
    If txtCMC72.Tag Then
        If Len(Trim(txtCMC72.Text)) = txtCMC72.MaxLength Then SendKeys "{ENTER}"
    End If

End Sub


Private Sub txtCMC72_GotFocus()
    
    txtCMC72.SelStart = 0
    txtCMC72.SelLength = txtCMC72.MaxLength

End Sub


Private Sub txtCMC72_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then KeyCode = 0

End Sub

Private Sub txtCMC72_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC72) < 10 Then
            Beep
            MsgBox "Digite todos os números do segundo campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            GoTo Sair
        End If
        'Verifica o código do documento
        If Left(txtCMC72.Text, 3) <> "592" Then
            Beep
            MsgBox "Para Capa de OCT o segundo campo do CMC7 deve começar com o Nr. 592.", vbInformation + vbOKOnly, App.Title
            GoTo Sair
        End If
        txtCMC73.SetFocus
        
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
    End If

    Exit Sub
    
Sair:
    txtCMC72_GotFocus
    txtCMC72.SetFocus

End Sub

Private Sub txtCMC73_Change()
    
    If txtCMC73.Tag Then
        If Len(Trim(txtCMC73.Text)) = txtCMC73.MaxLength Then SendKeys "{ENTER}"
    End If

End Sub

Private Sub txtCMC73_GotFocus()
    
    txtCMC73.SelStart = 0
    txtCMC73.SelLength = txtCMC73.MaxLength

End Sub


Private Sub txtCMC73_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then KeyCode = 0

End Sub

Private Sub txtCMC73_KeyPress(KeyAscii As Integer)
    
    Dim iErroCMC7 As Integer
    
    InibirTeclaAlfa KeyAscii
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Len(txtCMC73) < 12 Then
            Beep
            MsgBox "Digite todos os números do terceiro campo CMC7.", vbExclamation + vbOKOnly, "CMC-7"
            GoTo SairCMC73
        End If
        
        If Mid(txtCMC73, 2, 3) <> "592" Then
            Beep
            MsgBox "O terceiro campo do CMC7 não confere como Capa de OCT.", vbInformation + vbOKOnly, App.Title
            GoTo SairCMC73
        End If
        
        If Not CMC7Ok(iErroCMC7) Then
            Beep
            MsgBox Switch(iErroCMC7 = 1, "Primeiro", iErroCMC7 = 2, "Segundo", iErroCMC7 = 3, "Terceiro") & _
                    " campo do CMC7 não confere!", vbExclamation + vbOKOnly, App.Title
            If iErroCMC7 = 1 Then
                txtCMC71.SetFocus: GoTo Sair
            End If
            If iErroCMC7 = 2 Then
                txtCMC72.SetFocus: GoTo Sair
            End If
            If iErroCMC7 = 3 Then
                GoTo SairCMC73
            End If
        Else
            'Finaliza digitação
            cmdConfirmar_Click
        End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
    End If

    Exit Sub

SairCMC73:
    txtCMC73_GotFocus
    txtCMC73.SetFocus
    
Sair:
End Sub


Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

Private Function ValidarDados()
Dim iPosErro As Integer
Dim sLeitura As String

    sLeitura = Trim(txtCMC71.Text + txtCMC72.Text + txtCMC73.Text)
    ValidarDados = False
    
    If sLeitura = "" Or _
        Len(Trim(txtCMC71.Text)) <> 8 Or _
        Len(Trim(txtCMC72.Text)) <> 10 Or _
        Len(Trim(txtCMC73.Text)) <> 12 Then
        MsgBox "Favor informar todos campos do CMC7 !", vbExclamation + vbOKOnly, App.Title
        txtCMC71_GotFocus
        txtCMC71.SetFocus
        Exit Function
    End If
    
    If Len(sLeitura) = 30 And _
        Mid(sLeitura, 1, 3) = "409" And _
        Mid(sLeitura, 20, 3) = "592" And _
        Mid(sLeitura, 9, 3) = "592" And _
        Mid((sLeitura), 12, 6) = Mid(sLeitura, 24, 6) Then
    
        'Verifica se Capa de Malote é válido
        If Not CMC7Ok(iPosErro) Then
            GoTo Err_ValidarDados
        End If
    Else
        GoTo Err_ValidarDados
    End If
    
    ValidarDados = True
    Exit Function
    
Err_ValidarDados:
    MsgBox "CMC7 não confere!", vbExclamation + vbOKOnly, App.Title
    txtCMC71_GotFocus
    txtCMC71.SetFocus
    
End Function

Private Function CMC7Ok(ByRef iCampoCMC7Erro As Integer) As Boolean
'-------------------------------------------------------------------------------------
'       Validar campos de digitação do CMC7
'
' Parâmetros:   iCampoCMC7Erro    - Número do campo CMC7 com erro de verificação
'
' Retorno:      True    - CMC7 ok
'               False   - CMC7 com erro de dígito
'-------------------------------------------------------------------------------------
    
    Dim sCMC7 As String, sCmc71 As String, sCmc72 As String
    Dim sCmc73 As String, svalor As String
    
    CMC7Ok = False: iCampoCMC7Erro = 0
    
    'Verifica o código do documento
    If Left(txtCMC72.Text, 3) <> "592" Then iCampoCMC7Erro = 1: GoTo Sair
    If Mid(txtCMC73, 2, 3) <> "592" Then iCampoCMC7Erro = 3: GoTo Sair
    
    sCMC7 = Format(txtCMC71, "00000000") + Format(txtCMC72, "0000000000") + Format(txtCMC73, "000000000000")
    If Not TratarCamposCMC7(sCMC7, sCmc71, sCmc72, sCmc73, svalor) Then
        If Val(sCmc71) = 0 Then iCampoCMC7Erro = 1: GoTo Sair
        If Val(sCmc72) = 0 Then iCampoCMC7Erro = 2: GoTo Sair
        If Val(sCmc73) = 0 Then iCampoCMC7Erro = 3: GoTo Sair
    End If
    
    CMC7Ok = True

Sair:

End Function

Private Function AtualizaDocumentoExcluido(ByVal IdDocto As Long) As Boolean
    On Error GoTo ErroExclusao
    rdoErrors.Clear
    
    AtualizaDocumentoExcluido = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaDocumentoExcluido
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = "D" ' status
        .rdoParameters(4) = 1   ' duplicidade
        .rdoParameters(5) = 998 ' ocorrencia
        .Execute
        If .rdoParameters(0) <> 0 Then
            GoTo ErroExclusao
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroExclusao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Function

