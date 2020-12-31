VERSION 5.00
Object = "{39F894DF-E245-11D4-B08D-00600899AB13}#1.4#0"; "UbbLVImg.ocx"
Begin VB.Form frmIdentTipo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Identificação de Tipo"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10065
   StartUpPosition =   1  'CenterOwner
   Begin UbbLVImg.UBBImage imgIdent 
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7223
      ButEnabled3     =   0   'False
      ButEnabled4     =   0   'False
      ButEnabled5     =   0   'False
      ButEnabled6     =   0   'False
      ImageFile       =   "d:\tmp\get2.tif"
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Top             =   8100
      Width           =   1515
   End
   Begin UbbLVImg.UbbListView lvwTipoDoc 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   6906
      SortOnColumnClick=   0   'False
      MultiSelect     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumCols         =   1
      Col1            =   "Tipo de Documento"
      Tag1            =   "Capa de Bloco de Remetida Normal"
   End
   Begin UbbLVImg.UbbListView lvwBanco 
      Height          =   3915
      Left            =   3300
      TabIndex        =   1
      Top             =   4140
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   6906
      SortOnColumnClick=   0   'False
      UseIcons        =   0   'False
      MultiSelect     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumCols         =   2
      Col1            =   "Banco"
      Tag1            =   "000"
      Type1           =   "2"
      Col2            =   "Nome"
      Tag2            =   "Banco Bandeirantes  "
   End
   Begin UbbLVImg.UbbListView lvwTipif 
      Height          =   3915
      Left            =   6360
      TabIndex        =   2
      Top             =   4140
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   6906
      SortOnColumnClick=   0   'False
      UseIcons        =   0   'False
      MultiSelect     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumCols         =   2
      Col1            =   " "
      Tag1            =   "00"
      Type1           =   "1"
      Col2            =   "Tipificação"
      Tag2            =   "TB - Transferência Bancária (CPMF)"
   End
End
Attribute VB_Name = "frmIdentTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_blnTipif As Boolean


'Usuário finalizou a identificação
Private Sub cmdOk_Click()
    Hide
End Sub


Private Sub Form_Activate()
    'Acerta o foco...
    If lvwTipoDoc.Enabled And lvwTipoDoc.Visible Then
        lvwTipoDoc.SetFocus
    ElseIf lvwBanco.Enabled And lvwBanco.Visible Then
        lvwBanco.SetFocus
    ElseIf lvwTipif.Enabled And lvwTipif.Visible Then
        lvwTipif.SetFocus
    End If
    
    '...e a seleção
    If lvwTipoDoc.GetCount > 0 Then
        lvwTipoDoc.SelectItem 1
    End If
    If lvwBanco.GetCount > 0 Then
        lvwBanco.SelectItem 1
    End If
    If lvwTipif.GetCount > 0 Then
        lvwTipif.SelectItem 1
    End If
    
    lvwTipoDoc_Change
End Sub


'Trata o enter
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ActiveControl.Name = lvwTipoDoc.Name Then
            If lvwBanco.Enabled Then
                If lvwBanco.Visible Then lvwBanco.SetFocus
            ElseIf lvwTipif.Enabled Then
                If lvwTipif.Visible Then lvwTipif.SetFocus
            Else
                If cmdOk.Visible Then cmdOk.SetFocus
            End If
            KeyAscii = 0
        ElseIf ActiveControl.Name = lvwBanco.Name Then
            If lvwTipif.Enabled Then
                If lvwTipif.Visible Then lvwTipif.SetFocus
            Else
                If cmdOk.Visible Then cmdOk.SetFocus
            End If
            KeyAscii = 0
        ElseIf ActiveControl.Name = lvwTipif.Name Then
            If cmdOk.Visible Then cmdOk.SetFocus
            KeyAscii = 0
        End If
    End If
End Sub


'Recebe parâmetros
Public Sub SetParam(ByRef ftpClient As UbbFtp.UbbFtpRexec, _
                    ByRef cfgAmb As bcci.CConfig, _
                    ByRef cdbConn As bcci.CDatabase, _
                    ByVal lngSeq As Long, _
                    ByVal lngSeqMacroProc As Long, _
                    ByVal blnTipoDoc As Boolean, _
                    ByVal blnTipif As Boolean)
    Dim lngN As Long
    Dim vntArr As Variant
    
    'Obtem tipos de documentos válidos
    cdbConn.ExecSQL "select DESCRICAO, TIPO_DOC from TIPO_DOCUMENTO where " & _
                    "TIPO_DOC in (select TIPO_DOC from PROCESSO_TIPO_DOCUMENTO where SEQ_PROC in " & _
                    "(select SEQ_PROC from PROCESSO where SEQ_MACRO_PROC = ?)) order by TIPO_DOC", Array(lngSeqMacroProc), vntArr, lngN
    With lvwTipoDoc
        .Clear
        
        If lngN > 0 Then
            For lngN = 0 To UBound(vntArr, 2)
                .AddItem Array(vntArr(0, lngN), _
                               vntArr(1, lngN)), _
                         bcci.IconeTipoDoc(vntArr(1, lngN))
            Next
        End If
        
        .AddItem Array("Excluído", 0), icoExcluido
        
        .EndUpdate
    End With
    
    'Obtem bancos válidos
    cdbConn.ExecSQL "select SEQ_BANCO, NOME from MULTI_BANCO", , vntArr, lngN
    With lvwBanco
        .Clear
        
        If lngN > 0 Then
            For lngN = 0 To UBound(vntArr, 2)
                .AddItem Array(vntArr(0, lngN), _
                               vntArr(1, lngN)), _
                         icoNenhum
            Next
        End If
        
        .EndUpdate
    End With
    
    'Tipificação
    m_blnTipif = blnTipif
    With lvwTipif
        .Clear
        
        .AddItem Array(5, "Comum"), icoNenhum
        .AddItem Array(6, "OP - Ordem de Pagamento"), icoNenhum
        .AddItem Array(8, "ADM - Administrativo"), icoNenhum
        .AddItem Array(9, "TB - Transferência Bancária (CPMF)"), icoNenhum
        
        .EndUpdate
    End With
    
    'Carrega a imagem do documento a ser identificado
    Set imgIdent.UbbFtp = ftpClient
    imgIdent.DirLocal = cfgAmb.DirLocal
    imgIdent.DirBase = cfgAmb.DirBase
    imgIdent.DataMov = Mid$(cfgAmb.DataMov, 7, 4) & _
                       Mid$(cfgAmb.DataMov, 4, 2) & _
                       Mid$(cfgAmb.DataMov, 1, 2)
    
    imgIdent.LoadDoc lngSeq
    
    'Desabilita o que não será identificado
    lvwTipoDoc.Enabled = blnTipoDoc
    lvwBanco.Enabled = blnTipoDoc
    lvwTipif.Enabled = blnTipif

    Me.Show vbModal
End Sub


'Mudou o tipo de documento selecionado
Private Sub lvwTipoDoc_Change()
    Dim lngSel As Long
    Dim lngTipo As Long
    
    lngSel = lvwTipoDoc.SelectedLine
    
    If lngSel > 0 Then
        lngTipo = CLng(lvwTipoDoc.GetCel(2, lngSel))
        Select Case lngTipo
            Case 0 'Excluido
                lvwBanco.Enabled = False
                lvwTipif.Enabled = False
            Case bcci.TD_CHEQUE
                lvwBanco.Enabled = False
                If m_blnTipif Then
                    lvwTipif.Enabled = True
                End If
            Case Else
                lvwBanco.Enabled = True
                lvwTipif.Enabled = False
        End Select
    End If
End Sub


'Retorna tipo selecionado
Public Function GetTipoDoc() As Long
    Dim lngSel As Long
    
    lngSel = lvwTipoDoc.SelectedLine
    If (lngSel > 0) And (lvwTipoDoc.Enabled) Then
        GetTipoDoc = CLng(lvwTipoDoc.GetCel(2, lngSel))
    Else
        GetTipoDoc = 0
    End If
End Function


'Retorna banco selecionado
Public Function GetBanco() As Long
    Dim lngSel As Long
    
    lngSel = lvwBanco.SelectedLine
    If (lngSel > 0) And (lvwBanco.Enabled) Then
        GetBanco = CLng(lvwBanco.GetCel(1, lngSel))
    Else
        GetBanco = 0
    End If
End Function


'Retorna tipif selecionado
Public Function GetTipif() As Long
    Dim lngSel As Long
    
    lngSel = lvwTipif.SelectedLine
    If (lngSel > 0) And (lvwTipif.Enabled) Then
        GetTipif = CLng(lvwTipif.GetCel(1, lngSel))
    Else
        GetTipif = 0
    End If
End Function

