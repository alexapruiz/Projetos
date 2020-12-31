VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GrupoModulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo de M�dulos"
   ClientHeight    =   6384
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "M�dulos Selecionados"
      Height          =   4308
      Left            =   4560
      TabIndex        =   10
      Top             =   1032
      Width           =   3732
      Begin MSFlexGridLib.MSFlexGrid grdModuloSel 
         Height          =   3996
         Left            =   168
         TabIndex        =   11
         Top             =   216
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   7049
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "M�dulos Dispon�veis"
      Height          =   4308
      Left            =   120
      TabIndex        =   8
      Top             =   1032
      Width           =   3732
      Begin MSFlexGridLib.MSFlexGrid grdModulos 
         Height          =   3996
         Left            =   168
         TabIndex        =   9
         Top             =   216
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   7049
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton CmdRemoveTodos 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3960
      TabIndex        =   7
      ToolTipText     =   "Decrementa Todos"
      Top             =   3912
      Width           =   500
   End
   Begin VB.CommandButton CmdRemove1 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "Decrementa 1"
      Top             =   3264
      Width           =   500
   End
   Begin VB.CommandButton CmdIncrementaTodos 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Incrementa Todos"
      Top             =   2616
      Width           =   500
   End
   Begin VB.CommandButton cmdIncrementa1 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Incrementa 1"
      Top             =   1968
      Width           =   500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos"
      Height          =   756
      Left            =   1272
      TabIndex        =   2
      Top             =   96
      Width           =   6300
      Begin VB.ComboBox cboGrupos 
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   264
         Width           =   5604
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   708
      Left            =   7464
      Picture         =   "GrupoModulo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   840
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   6504
      Picture         =   "GrupoModulo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   816
   End
End
Attribute VB_Name = "GrupoModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* Querys *'
Dim qryGetModulos       As rdoQuery         '* Seleciona todos os M�dulos *'
Dim qryGetGrupos        As rdoQuery         '* Seleciona todos os Grupos  *'
Dim qryGetGrpModulo     As rdoQuery         '* Seleciona todos os Modulos do Grupo *'
Dim qryAlteraModulo     As rdoQuery         '* Altera Grupo M�dulo *'

'* RecordSet *'
Dim rsGetModulos        As rdoResultset     '* Recordset de M�dulos *'
Dim rsGetGrupos         As rdoResultset     '* Recordset de Grupos  *'
Dim rsGetGrpModulo      As rdoResultset     '* Recordset de M�dulos do Grupo *'
Private Sub EliminaModulosSelecionados()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' *        Elimina os M�dulos que fazem parte do Grupo selecionado no Combo,        * '
' *        para Grade de Modulos existentes no Sistema.                             * '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim Contador       As Integer '* Contador de Linhas da Grade de M�dulos selecionados *'
    Dim ContadorLs     As Integer '* Contador de Linhas da Grade de M�dulos existentes   *'
    Dim ContLinhasSel  As Integer '* Qtde de Linhas da Grade de M�dulos Selecionados     *'
    Dim ContLinhasMod  As Integer '* Qtde de Linhas da Grade de M�dulos do Sistema       *'

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Qtde de Linhas da Grade de M�dulos Selecionados * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ContLinhasSel = grdModuloSel.Rows - 1
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' *  Qtde de Linhas da Grade de M�dulos do Sistema  * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ContLinhasMod = grdModulos.Rows - 1

    '''''''''''''''''''''''''''''''''''
    ' *  Valor Inicial do Contador  * '
    '''''''''''''''''''''''''''''''''''
    For Contador = 1 To grdModuloSel.Rows - 1

        For ContadorLs = 1 To grdModulos.Rows - 1
            If grdModulos.TextMatrix(ContadorLs, 0) = grdModuloSel.TextMatrix(Contador, 0) Then
                If grdModulos.Rows = 2 Then
                    grdModulos.Rows = 1
                Else
                    grdModulos.RemoveItem (ContadorLs)
                    Exit For
                End If
            End If
        Next
    
    Next

Exit Sub
TrataErro:
    Select Case TratamentoErro("Erro ao eliminar M�dulos Selecionados.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Function GravaDados() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      * Cria Tratamento para Altera��o de Grupo M�dulo *                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim nControlaLoop  As Integer   '* Controlador de Loops *'
    Dim nContLinhas    As Integer   '* Conta todos as Linhas da Grade M�dulos Selecionados *'

    GravaDados = False

    Set qryAlteraModulo = Geral.Banco.CreateQuery("", "{Call AlteraGrupoModulo(?,?,?)}")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Qtde de Linhas que ser�o Incluidas na Base * '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    nContLinhas = grdModuloSel.Rows - 1
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Insere todos as linhas que est�o na Grade de M�dulos Selecionados * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For nControlaLoop = 1 To nContLinhas
    
        With qryAlteraModulo
            .rdoParameters(0) = cboGrupos.Text
            .rdoParameters(1) = CInt(grdModuloSel.TextMatrix(nControlaLoop, 1))
            .rdoParameters(2) = 2
            .Execute
        End With
            
    Next
    
    GravaDados = True
    
Exit Function
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel Gravar os Dados.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Function
Private Sub HabilitaBotoes()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'          * Habilita Bot�es de Incrementa��o de Decrementa��o de Itens das Grades *          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cmdIncrementa1.Enabled = True
    CmdIncrementaTodos.Enabled = True
    CmdRemove1.Enabled = True
    CmdRemoveTodos.Enabled = True

End Sub
Private Sub PreencheGrupoModulos()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         * Preenche Grade de Modulos Selecionados Para o Grupo Atual *            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim Contador As Integer

    Set qryGetGrpModulo = Geral.Banco.CreateQuery("", "{Call GetGrupoModulo(?)}")

    With qryGetGrpModulo
        ''''''''''''''''''''''''''
        ' * Descri��o do Grupo * '
        ''''''''''''''''''''''''''
        .rdoParameters(0) = cboGrupos.Text
        Set rsGetGrpModulo = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not rsGetGrpModulo.EOF Then
            
        '''''''''''''''''''''''''''''''''
        ' * Valor Inicial do Contador * '
        '''''''''''''''''''''''''''''''''
        Contador = 1
        
        ''''''''''''''''''''''''''''''''''''''''
        ' * Qtde de Linhas Iniciais da Grade * '
        ''''''''''''''''''''''''''''''''''''''''
        grdModuloSel.Rows = 1
        
        ''''''''''''''''''''''''''''''''''''
        ' * Qtde de Modulos Selecionados * '
        ''''''''''''''''''''''''''''''''''''
        grdModuloSel.Rows = rsGetGrpModulo.RowCount + 1
                
        Do While Not rsGetGrpModulo.EOF
            grdModuloSel.FocusRect = flexFocusNone
            grdModuloSel.TextMatrix(Contador, 0) = rsGetGrpModulo!Modulos
            grdModuloSel.TextMatrix(Contador, 1) = rsGetGrpModulo!IdModulo
            rsGetGrpModulo.MoveNext
            
            '''''''''''''''''''''''''''
            ' * Incrementa Contador * '
            '''''''''''''''''''''''''''
            Contador = Contador + 1
            
        Loop
    
    Else
        grdModuloSel.Clear
        grdModuloSel.Rows = 1
    End If

Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel preencher lista de M�dulos.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Function RemoveDados() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   * Deleta todos registros que pertencem ao Grupo Atual *                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    RemoveDados = False
    
    Set qryAlteraModulo = Geral.Banco.CreateQuery("", "{Call AlteraGrupoModulo(?,?,?)}")
        
    With qryAlteraModulo
        .rdoParameters(0) = cboGrupos.Text
        .rdoParameters(1) = 1
        .rdoParameters(2) = 1
        .Execute
    End With

    RemoveDados = True

Exit Function
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel excluir itens.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select
    
End Function

Private Sub cboGrupos_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      * Chama fun��o que lista todos os M�dulos do Grupo Selecionado nesta Grade *      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Len(Trim(cboGrupos.Text)) = 0 Then Exit Sub
    
        Screen.MousePointer = vbHourglass
        
            '''''''''''''''''''''''''''''''''
            ' * Habilita Todos os Bot�es  * '
            '''''''''''''''''''''''''''''''''
            Call HabilitaBotoes
        
            '''''''''''''''''''''''''''''''''
            ' * Preenche Grade de M�dulos * '
            '''''''''''''''''''''''''''''''''
            Call PreencheModulos
            
            ''''''''''''''''''''''''''''''''''''''''''
            ' * Preenche Grade de Grupo de M�dulos * '
            ''''''''''''''''''''''''''''''''''''''''''
            Call PreencheGrupoModulos
            
            ''''''''''''''''''''''''''''''''''''''''
            ' * Elimina Registros em Duplicidade * '
            ''''''''''''''''''''''''''''''''''''''''
            Call EliminaModulosSelecionados
    
        Screen.MousePointer = vbDefault
    
End Sub
Private Sub cmdConfirma_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      * Cria Tratamento para Altera��o de Grupo M�dulo *                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Len(Trim(cboGrupos.Text)) = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
        
        If RemoveDados Then
            If grdModuloSel.Rows > 1 Then
                GravaDados
            End If
        End If
        
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub cmdIncrementa1_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                        * Remove um registro da Grade de M�dulos  *                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
    
    With grdModulos
        
        If .Rows = 1 Then Exit Sub
        
            If .Rows = 2 Then
                .Rows = 1
                .FocusRect = flexFocusLight
                Exit Sub
            End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' * Incrementa 1 Na Grade de M�dulos Selecionados  * '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With grdModuloSel
            .Rows = .Rows + 1
            .FocusRect = flexFocusNone
            .TextMatrix(.Rows - 1, 0) = grdModulos.TextMatrix(grdModulos.Row, 0)
            .TextMatrix(.Rows - 1, 1) = grdModulos.TextMatrix(grdModulos.Row, 1)
            .SetFocus
            .Row = .Rows - 1
        End With
        
        '''''''''''''''''
        ' * Remove 1  * '
        '''''''''''''''''
        .RemoveItem (.Row)
    End With
    
Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel adicionar item selecionado.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Sub CmdIncrementaTodos_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    * Remove todos os registro da Grade de M�dulos *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    With grdModulos
        .FocusRect = flexFocusLight
        .Rows = 1
    End With
        
    Call PreencheModulosSel

Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel remover M�dulos.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Sub CmdRemove1_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   * Remove um registro da Grade de M�dulos Selecionados *                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    With grdModuloSel
        
        If .Rows = 1 Then Exit Sub
            If .Rows = 2 Then
                .Rows = 1
                .FocusRect = flexFocusLight
            Exit Sub
        End If
    
        ''''''''''''''''''''''''''''''''''''''''
        ' * Incrementa 1 na Grade de M�dulos * '
        ''''''''''''''''''''''''''''''''''''''''
        With grdModulos
            .Rows = .Rows + 1
            .FocusRect = flexFocusNone
            .TextMatrix(.Rows - 1, 0) = grdModuloSel.TextMatrix(grdModuloSel.Row, 0)
            .TextMatrix(.Rows - 1, 1) = grdModuloSel.TextMatrix(grdModuloSel.Row, 1)
            .SetFocus
            .Row = .Rows - 1
        End With
    
        ''''''''''''''''
        ' * Remove 1 * '
        ''''''''''''''''
        .RemoveItem (.Row)
    End With

Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel remover M�dulo selecionado.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Sub CmdRemoveTodos_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            * Remove todos os registro da Grade de M�dulos Selecionados *                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    With grdModuloSel
        .FocusRect = flexFocusLight
        .Rows = 1
    End With
        
    Call PreencheModulos
    
Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel remover M�dulos.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select
    
End Sub
Private Sub cmdSair_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   * Encerra Tela de Grupo de M�dulos *                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Unload Me
End Sub
Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       * Define Propriedades de inicializa��o *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    '''''''''''''''''''''''''''''''
    ' * Define Grade de M�dulos * '
    '''''''''''''''''''''''''''''''
    With grdModulos
        .Cols = 2
        .TextMatrix(0, 0) = String(22, " ") & "Descri��o do M�dulo"
        .ColWidth(0) = grdModulos.Width
        .ColWidth(1) = 0
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ' * Define Grade de M�dulos Selecionados * '
    ''''''''''''''''''''''''''''''''''''''''''''
    With grdModuloSel
        .Cols = 2
        .TextMatrix(0, 0) = String(22, " ") & "Descri��o do M�dulo"
        .ColWidth(0) = grdModulos.Width
        .ColWidth(1) = 0
    End With

    ''''''''''''''''''''''''''''''''
    ' * Preenche Combo de Grupos * '
    ''''''''''''''''''''''''''''''''
    Call PreencheGrupos
    
    '''''''''''''''''''''''''''''''''
    ' * Preenche Grade de M�dulos * '
    '''''''''''''''''''''''''''''''''
    Call PreencheModulos
    
Exit Sub
TrataErro:
    Select Case TratamentoErro("Erro ao inicializar M�dulo.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select
    
End Sub
Private Sub PreencheModulos()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        * Preenche Grade com todos os M�dulos cadastrados no Sistema *            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim Contador As Integer

    Set qryGetModulos = Geral.Banco.CreateQuery("", "{Call GetModulo}")
    Set rsGetModulos = qryGetModulos.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsGetModulos.EOF Then
    
        '''''''''''''''''''''''''''''''''''''''''''
        '      * Valor inicial do Contador *      '
        '''''''''''''''''''''''''''''''''''''''''''
        Contador = 1
    
        '''''''''''''''''''''''''''''''''''''''''''
        '    * Quantidade de Linhas Default *     '
        '''''''''''''''''''''''''''''''''''''''''''
        grdModulos.Rows = 1
              
        '''''''''''''''''''''''''''''''''''''''''''
        '    * Quantidade de Linhas da Grade *    '
        '''''''''''''''''''''''''''''''''''''''''''
        grdModulos.Rows = rsGetModulos.RowCount + 1
        
        Do While Not rsGetModulos.EOF
            grdModulos.FocusRect = flexFocusNone
            grdModulos.TextMatrix(Contador, 0) = rsGetModulos!Descricao
            grdModulos.TextMatrix(Contador, 1) = rsGetModulos!IdModulo
            
            '''''''''''''''''''''''''''''''''''''''''
            '      * Incrementa 1 ao contador *     '
            '''''''''''''''''''''''''''''''''''''''''
            Contador = Contador + 1
            rsGetModulos.MoveNext
        Loop
    
    End If

    '''''''''''''''''''''''''''
    '    * Linha Default *    '
    '''''''''''''''''''''''''''
    grdModulos.Row = 1
    
Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel preencher Grade de M�dulos.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select
    
End Sub
Private Sub PreencheModulosSel()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        * Preenche Grade com todos os M�dulos cadastrados no Sistema *            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim Contador As Integer

    Set qryGetModulos = Geral.Banco.CreateQuery("", "{Call GetModulo}")
    Set rsGetModulos = qryGetModulos.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsGetModulos.EOF Then
    
        '''''''''''''''''''''''''''''''''''''''''''
        '      * Valor inicial do Contador *      '
        '''''''''''''''''''''''''''''''''''''''''''
        Contador = 1
    
        '''''''''''''''''''''''''''''''''''''''''''
        '    * Quantidade de Linhas Default *     '
        '''''''''''''''''''''''''''''''''''''''''''
        grdModuloSel.Rows = 1
              
        '''''''''''''''''''''''''''''''''''''''''''
        '    * Quantidade de Linhas da Grade *    '
        '''''''''''''''''''''''''''''''''''''''''''
        grdModuloSel.Rows = rsGetModulos.RowCount + 1
        
        Do While Not rsGetModulos.EOF
            grdModuloSel.FocusRect = flexFocusNone
            grdModuloSel.TextMatrix(Contador, 0) = rsGetModulos!Descricao
            grdModuloSel.TextMatrix(Contador, 1) = rsGetModulos!IdModulo
            Contador = Contador + 1
            rsGetModulos.MoveNext
        Loop
    
    End If

Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel preencher Grade de M�dulos do Sistema.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Sub PreencheGrupos()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        * Preenche COMBO com todos os Grupos cadastrados no Sistema *            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Set qryGetGrupos = Geral.Banco.CreateQuery("", "{Call GetGrupos}")
    Set rsGetGrupos = qryGetGrupos.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsGetGrupos.EOF Then
    
        Do While Not rsGetGrupos.EOF
            cboGrupos.AddItem rsGetGrupos!Descricao
            rsGetGrupos.MoveNext
        Loop
    
    End If

Exit Sub
TrataErro:
    Select Case TratamentoErro("N�o foi poss�vel preencher lista de Grupos.", Err, rdoErrors)
        Case vbCancel
            cmdSair_Click
        Case vbRetry
    End Select

End Sub
Private Sub grdModulos_GotFocus()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            * Marca o Foco sempre na �ltima inclus�o feita na Grade de M�dulos *          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If grdModulos.Rows >= 18 Then
        With grdModulos
            .TopRow = .Row
        End With
    End If
End Sub
Private Sub grdModuloSel_GotFocus()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     * Marca o Foco sempre na �ltima inclus�o feita na Grade de M�dulos Selecionados *    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If grdModuloSel.Rows >= 18 Then
        With grdModuloSel
            .TopRow = .Row
        End With
    End If
End Sub
