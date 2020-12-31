VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmEstatistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acompanhamento de Produção"
   ClientHeight    =   7956
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5904
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7956
   ScaleWidth      =   5904
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   384
      Left            =   2976
      TabIndex        =   44
      Top             =   7500
      Width           =   1512
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   384
      Left            =   1416
      TabIndex        =   43
      Top             =   7500
      Width           =   1512
   End
   Begin VB.PictureBox Picture2 
      Height          =   252
      Left            =   2976
      ScaleHeight     =   204
      ScaleWidth      =   684
      TabIndex        =   41
      Top             =   528
      Width           =   732
      Begin VB.Label lblHora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "11:30:35"
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   252
      Left            =   2016
      ScaleHeight     =   204
      ScaleWidth      =   876
      TabIndex        =   39
      Top             =   528
      Width           =   924
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "17/07/2000"
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   -48
         TabIndex        =   40
         Top             =   0
         Width           =   936
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1740
      Left            =   144
      TabIndex        =   10
      Top             =   5568
      Width           =   5580
      Begin VB.Label lblPorCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   0
         Left            =   3504
         TabIndex        =   38
         Top             =   480
         Width           =   492
      End
      Begin VB.Label lblQtdDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.234.567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   0
         Left            =   4044
         TabIndex        =   37
         Top             =   480
         Width           =   876
      End
      Begin VB.Label lblPorDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   0
         Left            =   4944
         TabIndex        =   36
         Top             =   480
         Width           =   492
      End
      Begin VB.Label lblQtdCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "123.456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   0
         Left            =   2736
         TabIndex        =   35
         Top             =   480
         Width           =   732
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recepcionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   0
         Left            =   576
         TabIndex        =   34
         Top             =   480
         Width           =   2124
      End
      Begin VB.Label lblCor 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   264
         Index           =   0
         Left            =   156
         TabIndex        =   33
         Top             =   480
         Width           =   408
      End
      Begin VB.Label lblQtdCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "123.456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   1
         Left            =   2736
         TabIndex        =   32
         Top             =   780
         Width           =   732
      End
      Begin VB.Label lblPorDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   1
         Left            =   4944
         TabIndex        =   31
         Top             =   780
         Width           =   492
      End
      Begin VB.Label lblQtdDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.234.567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   1
         Left            =   4044
         TabIndex        =   30
         Top             =   780
         Width           =   876
      End
      Begin VB.Label lblPorCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   1
         Left            =   3504
         TabIndex        =   29
         Top             =   780
         Width           =   492
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Capturado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   1
         Left            =   588
         TabIndex        =   28
         Top             =   780
         Width           =   2124
      End
      Begin VB.Label lblCor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   264
         Index           =   1
         Left            =   144
         TabIndex        =   27
         Top             =   780
         Width           =   408
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   576
         TabIndex        =   26
         Top             =   192
         Width           =   2124
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   2736
         TabIndex        =   25
         Top             =   192
         Width           =   732
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   3504
         TabIndex        =   24
         Top             =   192
         Width           =   492
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doctos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4044
         TabIndex        =   23
         Top             =   192
         Width           =   876
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4944
         TabIndex        =   22
         Top             =   192
         Width           =   492
      End
      Begin VB.Label lblTotCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "123.456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   2736
         TabIndex        =   21
         Top             =   1392
         Width           =   732
      End
      Begin VB.Label lblTotPorCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   3504
         TabIndex        =   20
         Top             =   1392
         Width           =   492
      End
      Begin VB.Label lblTotDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.234.567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4044
         TabIndex        =   19
         Top             =   1392
         Width           =   876
      End
      Begin VB.Label lblTotPorDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4944
         TabIndex        =   18
         Top             =   1392
         Width           =   492
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   576
         TabIndex        =   17
         Top             =   1392
         Width           =   2124
      End
      Begin VB.Label lblCor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   264
         Index           =   2
         Left            =   144
         TabIndex        =   16
         Top             =   1068
         Width           =   408
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ocorrência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   2
         Left            =   576
         TabIndex        =   15
         Top             =   1056
         Width           =   2124
      End
      Begin VB.Label lblQtdCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "123.456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   2
         Left            =   2736
         TabIndex        =   14
         Top             =   1068
         Width           =   732
      End
      Begin VB.Label lblPorCapa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   2
         Left            =   3504
         TabIndex        =   13
         Top             =   1068
         Width           =   492
      End
      Begin VB.Label lblQtdDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.234.567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   2
         Left            =   4044
         TabIndex        =   12
         Top             =   1068
         Width           =   876
      End
      Begin VB.Label lblPorDoc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Index           =   2
         Left            =   4944
         TabIndex        =   11
         Top             =   1068
         Width           =   492
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   300
      Left            =   3168
      ScaleHeight     =   252
      ScaleWidth      =   1092
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   168
      Width           =   1140
      Begin VB.Label lblDataProc 
         Caption         =   "17/07/2000"
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
         Height          =   204
         Left            =   48
         TabIndex        =   8
         Top             =   24
         Width           =   1032
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   300
      Left            =   1284
      ScaleHeight     =   252
      ScaleWidth      =   1824
      TabIndex        =   5
      Top             =   168
      Width           =   1872
      Begin VB.Label Label8 
         Caption         =   "Data do Movimento"
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
         Height          =   192
         Left            =   24
         TabIndex        =   6
         Top             =   24
         Width           =   1740
      End
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   288
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Height          =   528
      Left            =   144
      TabIndex        =   4
      Top             =   5004
      Width           =   5580
      Begin VB.OptionButton optFiltro 
         Caption         =   "Malotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3540
         TabIndex        =   2
         Top             =   204
         Width           =   1236
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2112
         TabIndex        =   1
         Top             =   204
         Width           =   1236
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   984
         TabIndex        =   0
         Top             =   204
         Width           =   1236
      End
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   5004
      Left            =   192
      OleObjectBlob   =   "Estatistica.frx":0000
      TabIndex        =   9
      Top             =   240
      Width           =   5292
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7356
      Left            =   48
      TabIndex        =   3
      Top             =   48
      Width           =   5772
   End
End
Attribute VB_Name = "frmEstatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Connection        As rdo.rdoConnection
Public m_DataProcessamento As Long
Public m_Atualizacao       As Integer

Private Matriz(1 To 3, 1 To 2) As Long

Private qryGetEstatistica As rdoQuery
Private qryGetTotalCapa As rdoQuery
Private qryGetTotalDocumento As rdoQuery
Private rsEstatistica As rdoResultset
Private rsTotalCapa As rdoResultset
Private rsTotalDoc As rdoResultset

Private Function FormataQuantidade(ByVal Qtd As Long) As String
    Dim strValor As String
    Dim strResult As String
    Dim Count As Integer
    
    strValor = Trim(str(Qtd))
    
    For Count = 1 To Len(strValor)
        strResult = Mid(strValor, Len(strValor) - Count + 1, 1) & strResult
        If (Count Mod 3 = 0) And (Count < Len(strValor)) Then
            strResult = "." & strResult
        End If
    Next
    If Len(strResult) = 0 Then
        strResult = "0"
    End If
    FormataQuantidade = strResult
End Function

Private Function ObtemTotalCapa(ByVal IdEnv_Mal As String) As Long
    On Error GoTo ErroTotalCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetTotalCapa.rdoParameters(0) = m_DataProcessamento
    qryGetTotalCapa.rdoParameters(1) = IdEnv_Mal
    Set rsTotalCapa = qryGetTotalCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsTotalCapa.EOF Then
        ObtemTotalCapa = 0
    Else
        ObtemTotalCapa = rsTotalCapa!Total
    End If
    rsTotalCapa.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroTotalCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Erro na obtenção do total de Envelopes/Malotes.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemTotalCapa = -1

End Function

Private Function ObtemTotalDocumento(ByVal IdEnv_Mal As String) As Long
    On Error GoTo ErroTotalDoc
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetTotalDocumento.rdoParameters(0) = m_DataProcessamento
    qryGetTotalDocumento.rdoParameters(1) = IdEnv_Mal
    Set rsTotalDoc = qryGetTotalDocumento.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsTotalDoc.EOF Then
        ObtemTotalDocumento = 0
    Else
        ObtemTotalDocumento = rsTotalDoc!Total
    End If
    rsTotalDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroTotalDoc:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Erro na obtenção do total de Documentos.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemTotalDocumento = -1

End Function

Private Function ObtemEstatistica(ByVal IdEnv_Mal As String) As Boolean
    On Error GoTo ErroEstat
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetEstatistica.rdoParameters(0) = m_DataProcessamento
    qryGetEstatistica.rdoParameters(1) = IdEnv_Mal
    Set rsEstatistica = qryGetEstatistica.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    ObtemEstatistica = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroEstat:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Erro na obtenção da Estatística de Envelopes/Malotes.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    ObtemEstatistica = False

End Function

Private Sub Zera_Matriz()
    Dim Count As Integer
    
    For Count = 1 To 3
        Matriz(Count, 1) = 0
        Matriz(Count, 2) = 0
    Next
    
End Sub

Private Sub Preenche_Matriz()
    
    Zera_Matriz
    
    If rsEstatistica.RowCount > 0 Then
        rsEstatistica.MoveFirst
    End If
    
    While Not rsEstatistica.EOF
        Select Case rsEstatistica!Status
            Case "0" 'Recepcionado
                Matriz(1, 1) = Matriz(1, 1) + rsEstatistica!QtdCapa
                Matriz(1, 2) = Matriz(1, 2) + rsEstatistica!QtdDoc
            Case "1" 'Capturado
                Matriz(2, 1) = Matriz(2, 1) + rsEstatistica!QtdCapa
                Matriz(2, 2) = Matriz(2, 2) + rsEstatistica!QtdDoc
            Case "P", "D", "F", "X" 'Excluido / Ocorrencia
                Matriz(3, 1) = Matriz(3, 1) + rsEstatistica!QtdCapa
                Matriz(3, 2) = Matriz(3, 2) + rsEstatistica!QtdDoc
        End Select
        rsEstatistica.MoveNext
    Wend
    
    rsEstatistica.Close
End Sub

Private Sub AtualizaGrafico(ByVal IdEnv_Mal As String)
    Dim CountCapa As Long
    Dim CountDocto As Long
    Dim Count As Integer
    Dim TotPorCapa, TotPorDoc As Double
    
    tmrAtualiza.Enabled = False

    lblData.Caption = Format(Now, "dd/mm/yyyy")
    lblHora.Caption = Format(Now, "hh:mm:ss")
    
    CountCapa = ObtemTotalCapa(IdEnv_Mal)
    If CountCapa = -1 Then
        Exit Sub
    End If
    
    CountDocto = ObtemTotalDocumento(IdEnv_Mal)
    If CountDocto = -1 Then
        Exit Sub
    End If
    
    If Not ObtemEstatistica(IdEnv_Mal) Then
        Exit Sub
    End If
    
    Preenche_Matriz
    
    Grafico.Visible = False
    
    TotPorCapa = 0
    TotPorDoc = 0
    
    Grafico.ColumnCount = 3
    For Count = 1 To 3
        lblQtdCapa(Count - 1).Caption = FormataQuantidade(Matriz(Count, 1))
        lblQtdDoc(Count - 1).Caption = FormataQuantidade(Matriz(Count, 2))
        
        Grafico.Column = Count
        If CountCapa > 0 Then
            lblPorCapa(Count - 1).Caption = Format((Matriz(Count, 1) * 100 / CountCapa), "0.0")
            TotPorCapa = TotPorCapa + (Matriz(Count, 1) * 100 / CountCapa)
            Grafico.Data = (Matriz(Count, 1) * 100 / CountCapa)
        Else
            lblPorCapa(Count - 1).Caption = "0.0"
            Grafico.Data = 0
        End If
        
        If CountDocto > 0 Then
            lblPorDoc(Count - 1).Caption = Format((Matriz(Count, 2) * 100 / CountDocto), "0.0")
            TotPorDoc = TotPorDoc + (Matriz(Count, 2) * 100 / CountDocto)
        Else
            lblPorDoc(Count - 1).Caption = "0.0"
        End If
    Next
    
    lblTotCapa.Caption = FormataQuantidade(CountCapa)
    lblTotPorCapa.Caption = Format(TotPorCapa, "0.0")
    lblTotDoc.Caption = FormataQuantidade(CountDocto)
    lblTotPorDoc.Caption = Format(TotPorDoc, "0.0")
    
    Grafico.Visible = True
    
    tmrAtualiza.Enabled = True

End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    On Error GoTo ERRO_IMPRESSAO
    PrintForm
    On Error GoTo 0
    Exit Sub
ERRO_IMPRESSAO:
    MsgBox "Erro na impressão da Estatística.", vbCritical + vbOKOnly, App.Title
End Sub

Private Sub Form_Activate()
    tmrAtualiza.Interval = m_Atualizacao * 1000
    tmrAtualiza.Enabled = True
    lblDataProc.Caption = Mid(m_DataProcessamento, 7, 2) & "/" & _
        Mid(m_DataProcessamento, 5, 2) & "/" & _
        Mid(m_DataProcessamento, 1, 4)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Count As Integer
    
    Grafico.ColumnCount = 3
    If KeyCode = vbKeyF2 Then
        For Count = 1 To 3
            Grafico.Column = Count
            Grafico.Data = 33.33
        Next
    End If
End Sub

Private Sub Form_Load()
    Set qryGetEstatistica = m_Connection.CreateQuery("", "{Call MDIAG_GetEstatistica (?,?)}")
    Set qryGetTotalCapa = m_Connection.CreateQuery("", "{Call MDIAG_GetTotalCapa (?,?)}")
    Set qryGetTotalDocumento = m_Connection.CreateQuery("", "{Call MDIAG_GetTotalDocumento (?,?)}")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAtualiza.Enabled = False
    
    qryGetEstatistica.Close
    qryGetTotalCapa.Close
    qryGetTotalDocumento.Close
End Sub

Private Sub lblCor_DblClick(Index As Integer)
    If Matriz(Index + 1, 1) = 0 Then
        Exit Sub
    End If
    
    tmrAtualiza.Enabled = False
    
    Load frmListaCapa
    
    Set frmListaCapa.m_Connection = m_Connection
    frmListaCapa.m_DataProcessamento = m_DataProcessamento
    
    If optFiltro(0).Value Then
        frmListaCapa.m_IdEnv_Mal = "T"
    ElseIf optFiltro(1).Value Then
        frmListaCapa.m_IdEnv_Mal = "E"
    Else
        frmListaCapa.m_IdEnv_Mal = "M"
    End If
    
    Select Case Index
        Case 0
            frmListaCapa.m_InStatus = "'0'"
        Case 1
            frmListaCapa.m_InStatus = "'1'"
        Case 2
            frmListaCapa.m_InStatus = "'D','F','P','X'"
    End Select
    
    frmListaCapa.Caption = lblStatus(Index).Caption
    frmListaCapa.Show vbModal, Me
    
    tmrAtualiza.Enabled = True
    
End Sub

Private Sub optFiltro_Click(Index As Integer)
    If optFiltro(0).Value Then
        lblFiltro.Caption = "Todos"
        AtualizaGrafico ("T")
    ElseIf optFiltro(1).Value Then
        lblFiltro.Caption = "Env."
        AtualizaGrafico ("E")
    Else
        lblFiltro.Caption = "Malotes"
        AtualizaGrafico ("M")
    End If
End Sub

Private Sub tmrAtualiza_Timer()
    If optFiltro(0).Value Then
        AtualizaGrafico ("T")
    ElseIf optFiltro(1).Value Then
        AtualizaGrafico ("E")
    Else
        AtualizaGrafico ("M")
    End If
End Sub
