VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Relatorios 
   Caption         =   "SGB - Galeria de Relatórios"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   1035
      Top             =   3105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3735
      TabIndex        =   0
      Top             =   6210
      Width           =   1635
   End
End
Attribute VB_Name = "relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChequesPorData()

    Dim Rec As New ADODB.Recordset
    Dim sSql As String

    sSql = " SELECT CNT.ID_CNT AS CONTRATO, CNT.DATA_FESTA AS DATA_FESTA, CLI.NOM_CLI AS CLIENTE, PAR.DATA_PAR AS DATA_PARCELA , PAR.VALOR_PAR AS VALOR"
    sSql = sSql & " FROM CONTRATOS AS CNT, CLIENTES AS CLI, PARCELA_CONTRATO AS PAR "
    sSql = sSql & " WHERE CNT.ID_CNT = PAR.ID_CNT "
    sSql = sSql & " And CNT.COD_CLI = CLI.COD_CLI "
    sSql = sSql & " AND PAR.DATA_PAR BETWEEN #10/01/2006# AND #10/30/2006#"
    sSql = sSql & " ORDER BY PAR.DATA_PAR , CNT.DATA_FESTA"

    Rec.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Call Preenchegrid(Rec)
End Sub
Private Sub Preenchegrid(ByRef Rec As ADODB.Recordset)

    GridRel.Rows = 1
    GridRel.Cols = 0

    GridRel.Row = 0
    For x = 0 To Rec.Fields.Count - 1
        GridRel.Cols = GridRel.Cols + 1
        GridRel.Col = x
        GridRel.Text = Rec(x).Name

        GridRel.ColWidth(x) = 1200
    Next x

    GridRel.Rows = 2
    GridRel.Row = 1
    Do Until Rec.EOF
        GridRel.Rows = GridRel.Rows + 1
        For x = 0 To Rec.Fields.Count - 1
            GridRel.Col = x
            GridRel.Text = Rec(x).Value
        Next x
        GridRel.Row = GridRel.Row + 1
        Rec.MoveNext
    Loop
End Sub
Private Sub Command1_Click()

    CrystalReport.ReportFileName = App.Path & "\Relatorios\ChequesaDepositar.rpt"
    'CrystalReport.SelectionFormula = "{AvisoDiferenca.DataOcorrencia} = " & Geral.DataProcessamento & " and {AvisoDiferenca.Gerado} = " & Val(CBool(0))
    CrystalReport.CopiesToPrinter = 1
    'CrystalReport.WindowState = crptMaximized
    'CrystalReport.WindowTitle = "Relatório"
    CrystalReport.Action = 0
    
End Sub

