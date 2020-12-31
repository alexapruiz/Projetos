VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MensagemRobo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensagens do Robô"
   ClientHeight    =   3120
   ClientLeft      =   468
   ClientTop       =   3336
   ClientWidth     =   11256
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   11256
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBotoes 
      Height          =   1692
      Left            =   9168
      TabIndex        =   2
      Top             =   48
      Width           =   1932
      Begin VB.CommandButton cmdApagarMsg 
         Caption         =   "Apagar &Mensagem"
         Height          =   348
         Left            =   192
         TabIndex        =   5
         Top             =   672
         Width           =   1572
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   348
         Left            =   192
         TabIndex        =   4
         Top             =   1200
         Width           =   1572
      End
      Begin VB.CommandButton cmdAtualiza 
         Caption         =   "&Atualizar"
         Height          =   348
         Left            =   192
         TabIndex        =   3
         Top             =   240
         Width           =   1572
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdMensagemRobo 
      Height          =   2604
      Left            =   48
      TabIndex        =   1
      Top             =   144
      Width           =   8988
      _ExtentX        =   15854
      _ExtentY        =   4593
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   2844
      Width           =   11256
      _ExtentX        =   19854
      _ExtentY        =   487
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15092
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "15/5/2003"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "11:46"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MensagemRobo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_RstMensagem As rdoResultset

Private Sub ShowFields()

Dim i As Integer
    
    With grdMensagemRobo
        .TextMatrix(0, 0) = "Estação":      .ColAlignment(0) = flexAlignCenterCenter: .ColWidth(0) = .Width * 0.08
        .TextMatrix(0, 1) = "Data/Hora":    .ColAlignment(1) = flexAlignCenterCenter: .ColWidth(1) = .Width * 0.17
        .TextMatrix(0, 2) = "Capa":         .ColAlignment(2) = flexAlignCenterCenter: .ColWidth(2) = .Width * 0.15
        .Row = 0: .Col = 3
        .TextMatrix(0, 3) = "Descrição ":   .CellAlignment = flexAlignCenterCenter:   .ColWidth(3) = .Width * 0.6
    End With
    
    For i = 1 To m_RstMensagem.RowCount
        With grdMensagemRobo
            .TextMatrix(i, 0) = m_RstMensagem!IdEstacao
            .TextMatrix(i, 1) = Format(m_RstMensagem!DataHora, "dd/mm/yyyy HH:MM:SS")
            .TextMatrix(i, 2) = m_RstMensagem!Capa
            .Col = 3: .Row = i
            .TextMatrix(i, 3) = m_RstMensagem!Descricao: .CellAlignment = flexAlignLeftCenter
            .RowData(i) = m_RstMensagem!IdMensagemRobo  'Referência da mensagem (ID)
        End With
        
        m_RstMensagem.MoveNext
    Next
    
    'Apresenta linha de seleção
    grdMensagemRobo.Row = 1
    grdMensagemRobo.Col = 0
    grdMensagemRobo.ColSel = 3
    
    cmdApagarMsg.Enabled = True
    
End Sub

Public Function ShowModal(ByRef pRst As RDO.rdoResultset) As Boolean

    Set m_RstMensagem = pRst
    
    grdMensagemRobo.Rows = m_RstMensagem.RowCount + 1
    grdMensagemRobo.Cols = 4
   
    ShowFields

    MensagemRobo.Show vbModal

End Function

Private Sub cmdApagarMsg_Click()

Dim qryInsereMensagemRoboUsuario As rdoQuery
Dim i As Integer, iStep As Integer

On Error GoTo Err_cmdApagarMsg

    If grdMensagemRobo.Row <= grdMensagemRobo.RowSel Then
        iStep = 1
    Else
        iStep = -1
    End If

    Set qryInsereMensagemRoboUsuario = Geral.Banco.CreateQuery("", "{? = call InsereMensagemRoboUsuario(?,?,?)}")
    With qryInsereMensagemRoboUsuario
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(3) = Geral.idUsuario
    End With
    
    fraBotoes.Enabled = False
    Screen.MousePointer = vbHourglass
    
    For i = grdMensagemRobo.Row To grdMensagemRobo.RowSel Step iStep

        qryInsereMensagemRoboUsuario.rdoParameters(2) = grdMensagemRobo.RowData(i)
        qryInsereMensagemRoboUsuario.Execute
        
        If qryInsereMensagemRoboUsuario.rdoParameters(0).Value <> 0 Then
            fraBotoes.Enabled = True
            Screen.MousePointer = vbDefault
            
            MsgBox "Erro ao atualizar as mensagens do Robô.", vbCritical
            CmdSair_Click
            Exit Sub
        End If
    Next

Exit_cmdApagarMsg:
    fraBotoes.Enabled = True
    Screen.MousePointer = vbDefault
    qryInsereMensagemRoboUsuario.Close
    LerMensagensRobo
    Exit Sub

Err_cmdApagarMsg:
    fraBotoes.Enabled = True
    Screen.MousePointer = vbDefault
    qryInsereMensagemRoboUsuario.Close
    MsgBox "Erro ao atualizar as mensagens do Robô.", vbCritical
    CmdSair_Click

End Sub

Private Sub cmdAtualiza_Click()
    
    Call LerMensagensRobo
    
End Sub

Private Sub CmdSair_Click()

    Unload Me

End Sub

Private Sub LerMensagensRobo()

Dim qryGetMensagemRobo As rdoQuery
    
    Set qryGetMensagemRobo = Geral.Banco.CreateQuery("", "{call GetMensagemRobo(?,?)}")
    '''''''''''''''''''
    'Configura a Query'
    '''''''''''''''''''
    qryGetMensagemRobo.rdoParameters(0) = Geral.DataProcessamento
    qryGetMensagemRobo.rdoParameters(1) = Geral.idUsuario
    
    '''''''''
    'Executa'
    '''''''''
    Set m_RstMensagem = qryGetMensagemRobo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    ''''''''''''''''''''''''''''''
    'Se tem mensagem, abre a tela'
    ''''''''''''''''''''''''''''''
    If m_RstMensagem.EOF() Then
        grdMensagemRobo.Rows = 1
        cmdApagarMsg.Enabled = False
        MsgBox "Não há mensagem de retorno enviada pelo robô.", vbInformation, App.Title
        Exit Sub
    End If
    
    grdMensagemRobo.Rows = m_RstMensagem.RowCount + 1
    ShowFields

End Sub

