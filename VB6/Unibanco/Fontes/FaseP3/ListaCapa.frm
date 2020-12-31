VERSION 5.00
Begin VB.Form ListaCapa 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3792
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7092
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3792
   ScaleWidth      =   7092
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   372
      Left            =   2172
      TabIndex        =   1
      Top             =   3300
      Width           =   1488
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   3696
      TabIndex        =   2
      Top             =   3300
      Width           =   1488
   End
   Begin VB.Frame Frame1 
      Height          =   3192
      Left            =   48
      TabIndex        =   3
      Top             =   0
      Width           =   7008
      Begin VB.ListBox lstCapa 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2448
         ItemData        =   "ListaCapa.frx":0000
         Left            =   120
         List            =   "ListaCapa.frx":0002
         TabIndex        =   0
         Top             =   528
         Width           =   6780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agência"
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
         Height          =   288
         Left            =   5820
         TabIndex        =   7
         Top             =   204
         Width           =   1044
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
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
         Height          =   288
         Left            =   4440
         TabIndex        =   6
         Top             =   204
         Width           =   1332
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número Malote"
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
         Height          =   288
         Left            =   2424
         TabIndex        =   5
         Top             =   204
         Width           =   1968
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envelope / Malote"
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
         Height          =   288
         Left            =   120
         TabIndex        =   4
         Top             =   204
         Width           =   2256
      End
   End
End
Attribute VB_Name = "ListaCapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_InStatus               As String   '-- Status Capa
Public m_IdEnv_Mal              As String   '-- Identificador Envelope/Malote
Public m_tpGrafico              As Integer  '-- Tipo de gráfico 'T' / 'E'
Public m_idtela                 As Integer  '-- Identificador de Tela

'--Query de Pesquisa de Capas - Acompanhamento de Expedição--'
Private qryAcompExpedicaoCapa   As rdoQuery '-- Lista Todas as Capas
Private qryAcompExpedicaoCapaE  As rdoQuery '-- Lista as Capas de Envelope
Private qryAcompExpedicaoCapaM  As rdoQuery '-- Lista as Capas de Malote

Private rsCapa As rdoResultset
Private Function ObtemCapa() As Boolean
    Dim Linha As String
    Dim Sql As String
    
    '-- 0 Acompanhamento de Produção
    '-- 1 Acompanhamento de Expedição
    
    If m_idtela = 0 Then
    
        Sql = "Select Capa, Num_Malote, IdLote, AgOrig, IdEnv_Mal " & _
              "From Capa (NOLOCK)" & _
              "Where DataProcessamento = " & Trim(str(Geral.DataProcessamento)) & " And "
        If m_IdEnv_Mal <> "T" Then
            Sql = Sql & "IdEnv_Mal = '" & m_IdEnv_Mal & "' And "
        End If
        Sql = Sql & "Status in (" & m_InStatus & ") " & _
                    "Order By Idlote,IdEnv_Mal, Capa, AgOrig "
        
        On Error GoTo ErroCapa
        rdoErrors.Clear
        
        Screen.MousePointer = vbHourglass
        lstCapa.Clear
        
        Set rsCapa = Geral.Banco.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        
    ElseIf m_idtela = 1 Then
        
        '-- Trata Pesquisa (Acompanhamento de Expedição) --'
        If m_IdEnv_Mal = "T" Then
            With qryAcompExpedicaoCapa
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = m_InStatus
                .rdoParameters(2).Value = m_tpGrafico
                Set rsCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With
        ElseIf m_IdEnv_Mal = "E" Then
            With qryAcompExpedicaoCapaE
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = m_InStatus
                .rdoParameters(2).Value = m_tpGrafico
                Set rsCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With
        Else
            With qryAcompExpedicaoCapaM
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = m_InStatus
                .rdoParameters(2).Value = m_tpGrafico
                Set rsCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With
        End If
    
    End If
    
    While Not rsCapa.EOF
        If (rsCapa!IdEnv_Mal = "E" Or rsCapa!IdEnv_Mal = "F") Then
            'Envelope Cinza e Envelope Fininvest
            Linha = Format(rsCapa!Capa, "0000000000")
            Linha = Linha & Space(26) & Format(rsCapa!IdLote, "0000-00000")
            Linha = Linha & Space(2) & Format(rsCapa!AgOrig, "0000")
        'ElseIf rsCapa!IdEnv_Mal = "F" Then
        '    'Envelope Fininvest
        '    Linha = Format(rsCapa!Capa, "0000000000")
        '    Linha = Linha & Space(26) & Format(rsCapa!IdLote, "0000-00000")
        '    Linha = Linha & Space(2) & Format(rsCapa!AgOrig, "0000")
        Else
            'Malote
            Linha = Format(rsCapa!Capa, "00000000000000")
            If Left(Trim(rsCapa!Num_Malote), 1) = "9" And Len(Trim(rsCapa!Num_Malote)) = 11 Then
                Linha = Linha & Space(7) & Format(rsCapa!Num_Malote, "000000000000")
            Else
                Linha = Linha & Space(7) & Format(rsCapa!Num_Malote, "000000000000")
            End If
            Linha = Linha & Space(3) & Format(rsCapa!IdLote, "0000-00000")
            Linha = Linha & Space(2) & Format(rsCapa!AgOrig, "0000")
        End If
        lstCapa.AddItem Linha
        rsCapa.MoveNext
    Wend
    rsCapa.Close
    ObtemCapa = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção dos Envelopes/Malotes.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemCapa = False
    
End Function
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    On Error GoTo ERRO_IMPRESSAO

    Dim sColuna1 As String
    Dim iInd As Integer
    Dim iQtde As Integer
    
    Screen.MousePointer = vbHourglass
    
    iQtde = 80
    Printer.ScaleMode = 3
    Printer.Orientation = vbPRORPortrait
    Printer.Font = "Courier New"
    Printer.FontSize = 11
    
    For iInd = 0 To lstCapa.ListCount - 1
        iQtde = iQtde + 1
        If iQtde > 80 Then
            iQtde = 1
            If iInd > 1 Then
                Printer.NewPage
            End If
            Printer.Print Me.Caption
            Printer.Print " "
            sColuna1 = "Envelope/Malote" & Space(5) & _
                "Nr. Malote " & Space(5) & "Lote " & _
                Space(2) & "Agencia"
            Printer.Print sColuna1
            Printer.Print String(50, "-")
            Printer.Print " "
        End If
        Printer.Print lstCapa.List(iInd)
    Next iInd
    
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    DoEvents

    Exit Sub

ERRO_IMPRESSAO:
    MsgBox "Verifique se a impressora está conectada.", vbCritical + vbOKOnly, App.Title
End Sub
Private Sub Form_Activate()
    If Len(Trim(m_InStatus)) > 0 And Len(Trim(m_IdEnv_Mal)) > 0 Then
        If ObtemCapa Then
            If lstCapa.ListCount > 0 Then
                lstCapa.Selected(0) = True
            End If
        End If
    End If
End Sub
Private Sub Form_Load()

    '-- Pesquisa de Capas Acompanhamento de Expedicao--'
    Set qryAcompExpedicaoCapa = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicaoCapa(?,?,?)}")

    '-- Pesquisa de Capas Acompanhamento de Expedicao--'
    Set qryAcompExpedicaoCapaE = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicaoCapaE(?,?,?)}")

    '-- Pesquisa de Capas Acompanhamento de Expedicao--'
    Set qryAcompExpedicaoCapaM = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicaoCapaM(?,?,?)}")

End Sub
