VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form Impressao 
   Caption         =   "Form1"
   ClientHeight    =   5556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5808
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5556
   ScaleWidth      =   5808
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim m_Report As Object

Public Enum enumTipoImpressao
    eBorderoGerenciamentoCheques
    eAvisoDiferenca
End Enum
Public Sub PrintReport(ByVal pPrintPreview As Boolean)

    m_Report.DiscardSavedData
    If pPrintPreview Then
        
        Me.Show vbModal
    Else
        m_Report.PrintOut False
    End If

End Sub

Public Sub SetImpressao(ByVal pTipoImpressao As enumTipoImpressao)

    If pTipoImpressao = eBorderoGerenciamentoCheques Then
        Set m_Report = New rptBorderoGerenciamentoCheques
        Me.Caption = "Borderô de Gerenciamento de Cheques"
    ElseIf pTipoImpressao = eAvisoDiferenca Then
        Set m_Report = New rptAvisoDiferenca
        Me.Caption = "Aviso de Diferênça"
    End If

End Sub

Public Sub SetSelectionFormula(ByVal pSelectionFormula As String)


    m_Report.DiscardSavedData

    ''''''''''''''''''''''''''''''''''''''''''
    'Acerta a Formula de Selecao do Relatorio'
    ''''''''''''''''''''''''''''''''''''''''''
    m_Report.GroupSelectionFormula = pSelectionFormula
    
    m_Report.ReadRecords
    
    CRViewer1.Refresh

End Sub

Private Sub Form_Load()

    

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault


    


End Sub


Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub


