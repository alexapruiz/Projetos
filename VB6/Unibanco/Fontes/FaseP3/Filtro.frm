VERSION 5.00
Begin VB.Form Filtro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   2616
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   5232
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2616
   ScaleWidth      =   5232
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTitulo 
      Caption         =   "Filtrar por títulos"
      Height          =   216
      Left            =   216
      TabIndex        =   3
      Top             =   72
      Value           =   -1  'True
      Width           =   1452
   End
   Begin VB.OptionButton optModulo 
      Caption         =   "Filtrar por módulo"
      Height          =   216
      Left            =   216
      TabIndex        =   2
      Top             =   768
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3864
      TabIndex        =   1
      Top             =   636
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3864
      TabIndex        =   0
      Top             =   156
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1764
      Left            =   120
      TabIndex        =   4
      Top             =   744
      Width           =   3588
      Begin VB.ListBox lstModulo 
         Height          =   1392
         Left            =   144
         TabIndex        =   5
         Top             =   288
         Width           =   3156
      End
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   3588
      Begin VB.ListBox lstTitulo 
         Height          =   240
         Left            =   144
         TabIndex        =   7
         Top             =   264
         Width           =   3156
      End
   End
End
Attribute VB_Name = "Filtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim m_IdItem        As Integer
Dim m_Opcao         As Integer
Dim m_Ok            As Integer
Public Function ShowModal(ByRef pIdItem As Integer) As Integer

    Dim i       As Integer

    If pIdItem <> 0 Then
        For i = 0 To lstModulo.ListCount - 1
            If lstModulo.ItemData(i) = pIdItem Then
                lstModulo.Selected(i) = True
                Exit For
            End If
        Next i
    End If
    
    Filtro.Show vbModal
    
    ''''''''''''''''''''''''''''
    '0 - Cancelou              '
    '1 - Selecionou Titulo     '
    '2 - Selecionou Modulo     '
    ''''''''''''''''''''''''''''
    
    ShowModal = m_Ok
    pIdItem = m_IdItem

End Function

Private Sub cmdCancelar_Click()
    m_Ok = 0
    m_IdItem = 0
    
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim i           As Integer

    m_Ok = IIf(optTitulo.Value = True, 1, 2)
    m_IdItem = 0
    
    Select Case m_Ok
    Case 1
        '''''''''''''''
        'Se for Título'
        '''''''''''''''
        For i = 0 To lstTitulo.ListCount - 1
            If lstTitulo.Selected(i) = True Then
                m_IdItem = lstTitulo.ItemData(i)
                Exit For
            End If
        Next i
        If m_IdItem = 0 Then
            MsgBox "Selecione um título.", vbExclamation
            m_Ok = 0
            Exit Sub
        End If
    Case 2
        '''''''''''''''
        'Se for módulo'
        '''''''''''''''
        For i = 0 To lstModulo.ListCount - 1
            If lstModulo.Selected(i) = True Then
                m_IdItem = lstModulo.ItemData(i)
                Exit For
            End If
        Next i
        If m_IdItem = 0 Then
            MsgBox "Selecione um módulo.", vbExclamation
            m_Ok = 0
            Exit Sub
        End If
    End Select
    Unload Me
    
End Sub


Private Sub Form_Load()

    Dim qryGetModuloFiltro      As RDO.rdoQuery
    Dim rst                     As RDO.rdoResultset
    
    ''''''''''''''''''''''''''''
    'Preenche o List de Modulos'
    ''''''''''''''''''''''''''''
    Set qryGetModuloFiltro = Geral.Banco.CreateQuery("", "{Call GetModuloFiltro}")
    
    Set rst = qryGetModuloFiltro.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    Do While Not rst.EOF()
    
        lstModulo.AddItem rst!Descricao
        lstModulo.ItemData(lstModulo.NewIndex) = rst!IdModulo
    
        rst.MoveNext
    Loop
    
    ''''''''''''''''''''''''''''
    'Preenche o List de Titulos'
    ''''''''''''''''''''''''''''
    lstTitulo.AddItem "Títulos de outros bancos"
    lstTitulo.ItemData(lstTitulo.NewIndex) = 31 'titulo de outros bancos
    

End Sub


Private Sub lstModulo_Click()
    optModulo.Value = True
    
End Sub


Private Sub lstTitulo_Click()
    optTitulo.Value = True
End Sub


