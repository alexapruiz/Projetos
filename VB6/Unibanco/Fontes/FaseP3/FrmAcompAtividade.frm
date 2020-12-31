VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAcompAtividade 
   Caption         =   "Acompanhamento de Atividades"
   ClientHeight    =   6948
   ClientLeft      =   2016
   ClientTop       =   1164
   ClientWidth     =   8712
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6948
   ScaleWidth      =   8712
   Begin VB.TextBox TxtTotal 
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
      Height          =   288
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5760
      Width           =   852
   End
   Begin VB.CommandButton CmdLocalizar 
      Caption         =   "&Localizar"
      Enabled         =   0   'False
      Height          =   288
      Left            =   7680
      TabIndex        =   3
      Top             =   5760
      Width           =   852
   End
   Begin VB.ComboBox CboUsuario 
      Height          =   288
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5760
      Width           =   3132
   End
   Begin VB.CommandButton CmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   4440
      TabIndex        =   1
      Top             =   6360
      Width           =   972
   End
   Begin VB.CommandButton CmdAtualizar 
      Caption         =   "&Atualizar"
      Height          =   372
      Left            =   3240
      TabIndex        =   0
      Top             =   6360
      Width           =   972
   End
   Begin VB.Frame FraUsuario 
      Height          =   5172
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   4092
      Begin MSFlexGridLib.MSFlexGrid GrdUsuario 
         Height          =   4812
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3852
         _ExtentX        =   6795
         _ExtentY        =   8488
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame FraModulo 
      Height          =   5172
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4092
      Begin MSFlexGridLib.MSFlexGrid GrdModulo 
         Height          =   4812
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   3852
         _ExtentX        =   6795
         _ExtentY        =   8488
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Total de Usuários :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   1200
      TabIndex        =   9
      Top             =   5760
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Usuários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   4440
      TabIndex        =   7
      Top             =   240
      Width           =   4092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Módulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4092
   End
End
Attribute VB_Name = "FrmAcompAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TpModulo                   '* Guarda Informação de Módulos *'
    IdModulo            As Integer      '- Identificação de Módulos
    NmModulo            As String * 50  '- Nome do Módulo
End Type

Private Type TpUsuario                  '* Guarda Informação de Usuários *'
    idUsuario           As Integer      '- Identificação dos Usuários
    NmUsuario           As String       '- Nome do Usuário
End Type

Private TpModulo()      As TpModulo
Private TpUsuario()     As TpUsuario

Private qryGetModulo    As rdoQuery     'Query GetModulos
Private qryGetUsuario   As rdoQuery     'Query GetUsuarios
Private qryGetAtividade As rdoQuery     'Query GetAcompAtividade

Private RsGetModulo     As rdoResultset 'Recordset GetModulos
Private RsGetUsuario    As rdoResultset 'Recordset GetUsuarios
Private RsGetAtividade  As rdoResultset 'Recordset GetAcompAtividade

Private Contador        As Long         'Variável Auxiliar de Contagem
Private Consulta        As Boolean

'* Usado para encontrar string no combo *'
Const CB_FINDSTRING = &H14C
Const CB_ERR = -1
Const CB_SETCURSEL = &H14E

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Function AutoMatch(prmCBO_ As Object, prmKeyAscii As Integer) As Long
'* Rotina de Busca no Combo de Usuários *'

On Error GoTo TrataErro

    Dim lclBuffer As String
    Dim lclRetVal As Long
    
    If prmCBO_.Locked Then Exit Function
    
    Err = 0
    
    lclBuffer = Left(prmCBO_.Text, prmCBO_.SelStart) & Chr(prmKeyAscii)
    lclRetVal = SendMessage((prmCBO_.hwnd), CB_FINDSTRING, -1, ByVal lclBuffer)
    
    If lclRetVal <> CB_ERR Then
        lclRetVal = SendMessage((prmCBO_.hwnd), CB_SETCURSEL, ByVal lclRetVal, -1)
        prmCBO_.Text = prmCBO_.List(lclRetVal)
        prmCBO_.SelStart = Len(lclBuffer)
        prmCBO_.SelLength = Len(prmCBO_.Text)
        CmdLocalizar.Enabled = True
        prmKeyAscii = 0
        Consulta = True
    Else
        CmdLocalizar.Enabled = False
        Consulta = False
    End If
    
    AutoMatch = lclRetVal

Exit Function

TrataErro:
    Select Case TratamentoErro("Erro ao buscar informação no Combo Usuário.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
    
End Function
Private Sub CboUsuario_Click()
'Permite click no botão localizar
    CmdLocalizar.Enabled = True
End Sub
Private Sub CboUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AutoMatch CboUsuario, KeyAscii
        CmdLocalizar_Click
    Else
        AutoMatch CboUsuario, KeyAscii
    End If
End Sub
Private Sub CmdAtualizar_Click()
'* Chama Rotinas para atualização da Grade de Módulos *'
    CboUsuario.Text = ""
    CmdLocalizar.Enabled = False
    Consulta = True
    Call AjustaGrade
    Call AtualizaModulos(0)
End Sub
Private Sub CmdFechar_Click()
    'Sai do Módulo
    Unload Me
End Sub
Private Sub CmdLocalizar_Click()
'* Chama rotina de localização de usuário *'
    If Len(Trim(CboUsuario.Text)) = 0 Then Exit Sub
    Call AjustaGrade
    If CboUsuario.ListIndex <= -1 Then
        Call AtualizaModulos(0)
        MsgBox "Usuário não cadastrado.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    Else
        Consulta = False
    End If
    Call AtualizaModulos(CboUsuario.ItemData(CboUsuario.ListIndex))
End Sub
Private Sub Form_Load()
'* Ajustes Iniciais *'
    Consulta = True
    Call AjustaGrade
    Call AtualizaUsuarios
    Call AtualizaModulos(0)
End Sub
Private Function AtualizaModulos(lidusuario As Long)
'* Objetivo: Atualizar Grade de Módulos *'
On Error GoTo TrataErro
    
    Erase TpModulo
    GrdModulo.Visible = False
    GrdModulo.Rows = 1
    
    '* Inicializa Querys *'
    Set qryGetModulo = Geral.Banco.CreateQuery("", "{Call GetModulos(?,?)}")

    With qryGetModulo
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = lidusuario
        .Execute
        Set RsGetModulo = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If RsGetModulo.EOF Then Exit Function
        
        'Redimensiona tamanho do Type TpModulo
        ReDim TpModulo(RsGetModulo.RowCount)
        
        'Redimensiona tamanho da Grade GrdModulo
        GrdModulo.Rows = RsGetModulo.RowCount + 1
        
        Do While Not RsGetModulo.EOF
            '* Preenche Type TpModulo *'
            TpModulo(Contador).IdModulo = RsGetModulo!IdModulo
            TpModulo(Contador).NmModulo = RsGetModulo!Descricao
            Contador = Contador + 1
            
            With GrdModulo
                '* Preenche Grade GrdModulo *'
                .TextMatrix(Contador, 0) = RsGetModulo!Descricao
                .ColAlignment(1) = 3
                .TextMatrix(Contador, 1) = RsGetModulo!Qtde
                .Col = 2
            End With
            
            RsGetModulo.MoveNext
        Loop
        
        RsGetModulo.MoreResults
            
        If Not RsGetModulo.EOF Then
            txtTotal = RsGetModulo!Sum
        End If
        
        qryGetModulo.Close
        Contador = Empty
        GrdModulo.Visible = True
        
Exit Function
TrataErro:
       GrdModulo.Visible = True
       qryGetModulo.Close
       Contador = Empty
    
    If Err = 40041 Then
       Call AjustaGrade
       AtualizaModulos (0)
       MsgBox "Usuário não está em atividade.", vbInformation + vbOKOnly
    Else
        Select Case TratamentoErro("Erro ao Atualizar Informação de Módulos.", Err, rdoErrors)
            Case vbCancel, vbRetry
                Unload Me
        End Select
    End If
        
End Function
Private Function AtualizaUsuarios()
'* Objetivo: Atualizar Informações de Usuários *'
    
On Error GoTo TrataErro
    
    Erase TpUsuario
    CboUsuario.Clear
    
    Set qryGetUsuario = Geral.Banco.CreateQuery("", "{Call GetUsuarios}")

    With qryGetUsuario
        .Execute
        Set RsGetUsuario = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If RsGetUsuario.EOF Then Exit Function
    
        'Redimensiona tamanho do Type TpUsuarios
        ReDim TpUsuario(RsGetUsuario.RowCount)
        
        Do While Not RsGetUsuario.EOF
            '* Preenche Type TpUsuario *'
            TpUsuario(Contador).idUsuario = RsGetUsuario!idUsuario
            TpUsuario(Contador).NmUsuario = RsGetUsuario!Nome
            Contador = Contador + 1
            
            With CboUsuario
                .AddItem RsGetUsuario!Nome
                .ItemData(CboUsuario.NewIndex) = RsGetUsuario!idUsuario
            End With
            
            RsGetUsuario.MoveNext
        Loop
        
        qryGetUsuario.Close
        Contador = Empty
Exit Function

TrataErro:
    qryGetUsuario.Close
    Contador = Empty
    
    Select Case TratamentoErro("Erro ao Atualizar Informação de Usuários.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
        
End Function
Private Sub GrdModulo_Click()
'* Atualização da Grade de Usuários *'
    If GrdModulo.Rows = 1 Then Exit Sub
    If Consulta = False Then Exit Sub
    Call AtualizaGrdUsuario
End Sub
Private Function AtualizaGrdUsuario()
'*  Objetivo: Atualizar Grade de Usuários   *'
    
On Error GoTo TrataErro
    
    GrdUsuario.Rows = 1
    
    Set qryGetAtividade = Geral.Banco.CreateQuery("", "{Call GetAcompAtividade(?,?)}")
    
    With qryGetAtividade
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = TpModulo(GrdModulo.Row - 1).IdModulo
        Set RsGetAtividade = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If RsGetAtividade.EOF Then Exit Function

        'Redimensiona tamanho da Grade de Usuários
        GrdUsuario.Rows = RsGetAtividade.RowCount + 1
        
        Do While Not RsGetAtividade.EOF
            Contador = Contador + 1
            
            With GrdUsuario
                '* Preenche Grade GrdUsuario *'
                .TextMatrix(Contador, 0) = RsGetAtividade!Nome
            End With
            
            RsGetAtividade.MoveNext
        Loop
        
        qryGetAtividade.Close
        Contador = Empty

Exit Function

TrataErro:
    qryGetUsuario.Close
    Contador = Empty
    
    Select Case TratamentoErro("Erro ao Atualizar Grade de Usuários.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select

End Function
Private Function AjustaGrade()

'*  Ajustes Iniciais Grade de Módulos   *'
    With GrdModulo
        .Visible = False
        .ColWidth(0) = .Width * 0.75
        .ColWidth(1) = .Width * 0.21
        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Módulos"
        .Col = 1
        .Text = "Qtde"
        .Col = 2
        .ColWidth(2) = .Width * 0
        .Visible = True
    End With

'*  Ajustes Iniciais Grade de Módulos   *'
    With GrdUsuario
        .ColWidth(0) = .Width * 1
        .ColWidth(1) = .Width * 0
        .ColWidth(2) = .Width * 0
        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Usuários"
    End With
    
    txtTotal.Text = ""
        
End Function
