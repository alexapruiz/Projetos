VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ListaContratos 
   Caption         =   "SGB - Lista de Contratos"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11940
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Ordenar por"
      Height          =   645
      Left            =   180
      TabIndex        =   13
      Top             =   1035
      Width           =   10095
      Begin VB.OptionButton OptDataContrato 
         Caption         =   "Data Contrato"
         Height          =   240
         Left            =   8415
         TabIndex        =   17
         Top             =   270
         Width           =   1410
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Código Cliente"
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   270
         Width           =   1410
      End
      Begin VB.OptionButton OptFesta 
         Caption         =   "Data Festa"
         Height          =   195
         Left            =   3765
         TabIndex        =   15
         Top             =   270
         Width           =   1230
      End
      Begin VB.OptionButton OptContrato 
         Caption         =   "Número Contrato"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   1680
      End
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   870
      Left            =   10485
      Picture         =   "ListaContratos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   495
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cliente"
      Height          =   870
      Left            =   5355
      TabIndex        =   12
      Top             =   135
      Width           =   4920
      Begin VB.ComboBox CboCliente 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   4560
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Número"
      Height          =   870
      Left            =   3465
      TabIndex        =   10
      Top             =   135
      Width           =   1770
      Begin VB.TextBox TxtContrato 
         Height          =   330
         Left            =   900
         TabIndex        =   0
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   870
      Left            =   180
      TabIndex        =   7
      Top             =   135
      Width           =   3165
      Begin VB.TextBox TxtAte 
         Height          =   330
         Left            =   2070
         TabIndex        =   2
         Top             =   315
         Width           =   1005
      End
      Begin VB.TextBox TxtDe 
         Height          =   330
         Left            =   495
         TabIndex        =   1
         Text            =   "01/01/2007"
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   240
         Left            =   1665
         TabIndex        =   9
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   210
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   420
      Left            =   4995
      TabIndex        =   5
      Top             =   6840
      Width           =   1950
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4875
      Left            =   135
      TabIndex        =   6
      Top             =   1755
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8599
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "ListaContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdPesquisar_Click()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String
    Dim TIPO_FESTA As String

    g.Rows = 1
    'Verificar qual critério de seleção o usuário utilizou
    If Len(Trim(TxtDe.Text)) > 0 And Len(Trim(TxtAte.Text)) > 0 Then
        'Verificar se foram informadas datas válidas
        If (Not IsDate(TxtDe.Text)) Or (Not IsDate(TxtAte.Text)) Then
            MsgBox "Datas informadas no período não são válidas.", vbOKOnly, "SGB"
            Exit Sub
        End If

        'Pesquisar contratos dentro do período informado
        sSql = "SELECT * FROM CONTRATOS , HORARIO_FESTA WHERE DATA_FESTA BETWEEN #" & Format(TxtDe.Text, "mm/dd/yyyy") & "# AND #" & Format(TxtAte.Text, "mm/dd/yyyy") & "#"

    ElseIf Val(TxtContrato.Text) > 0 Then

        sSql = " SELECT * FROM CONTRATOS ,  HORARIO_FESTA WHERE ID_CNT = " & TxtContrato.Text
    ElseIf CboCliente.ListIndex > -1 Then

        sSql = " SELECT * FROM CONTRATOS , HORARIO_FESTA WHERE COD_CLI = " & CboCliente.ItemData(CboCliente.ListIndex)
    Else
        MsgBox "Informe um critério para seleção !", vbOKOnly, "SGB"
        Exit Sub
    End If

    sSql = sSql & " AND HR_INI = ID_HORARIO "

    'Verificar o critério de ordenação selecionado pelo usuário
    If OptContrato.Value = True Then
        sSql = sSql & " ORDER BY ID_CNT "
    ElseIf OptCliente.Value = True Then
        sSql = sSql & " ORDER BY COD_CLI "
    ElseIf OptFesta.Value = True Then
        sSql = sSql & " ORDER BY DATA_FESTA , DSC_HORARIO "
    ElseIf OptDataContrato.Value = True Then
        sSql = sSql & " ORDER BY DATA_CNT "
    End If

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF Then

        x = 1
        Do Until Rs.EOF

            If Rs("TIPO_FESTA") = "C" Then TIPO_FESTA = "COMPLETA"
            If Rs("TIPO_FESTA") = "E" Then TIPO_FESTA = "ECONÔMICA"
            
            g.AddItem 0 & Chr(9) & _
                      Rs("ID_CNT").Value & Chr(9) & _
                      Rs("COD_CLI").Value & Chr(9) & _
                      Rs("NOME_ANIV").Value & Chr(9) & _
                      Rs("DATA_FESTA").Value & Chr(9) & _
                      Rs("DSC_HORARIO").Value & Chr(9) & _
                      TIPO_FESTA & Chr(9) & _
                      Rs("DATA_CNT").Value & Chr(9) & _
                      Rs("VALOR_TOTAL").Value & Chr(9) & _
                      Rs("DSC_BOLO").Value & Chr(9) & _
                      Rs("DSC_DECOR").Value & Chr(9) & _
                      Rs("OBS_BEBIDA").Value & Chr(9) & _
                      Rs("OBS").Value, x
            x = x + 1
            Rs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()

    Unload Me
End Sub
Private Sub Form_Load()

    Call CarregaCombo(CboCliente, "CLIENTES", "COD_CLI", "NOM_CLI", "")
    
    Call FormataGrid
    
End Sub
Private Sub FormataGrid()

    g.Cols = 13
    g.Rows = 2

    g.Row = 0

    g.Col = 1
    g.Text = "Contrato"

    g.Col = 2
    g.Text = "Cliente"

    g.Col = 3
    g.Text = "Aniv."
    
    g.Col = 4
    g.Text = "Data Festa"
    
    g.Col = 5
    g.Text = "Horário"

    g.Col = 6
    g.Text = "Tipo Festa"

    g.Col = 7
    g.Text = "Data Contrato"
    
    g.Col = 8
    g.Text = "Valor Total"
    
    g.Col = 9
    g.Text = "Bolo"
    
    g.Col = 10
    g.Text = "Decoração"
    
    g.Col = 11
    g.Text = "Bebida"
    
    g.Col = 12
    g.Text = "OBS"
    
    g.ColWidth(0) = 1
    g.ColAlignment(0) = 1

    g.ColWidth(1) = 800
    g.ColAlignment(1) = 3
    
    g.ColWidth(2) = 800
    g.ColAlignment(2) = 3
    
    g.ColWidth(3) = 2000
    g.ColAlignment(3) = 1
    
    g.ColWidth(4) = 1000
    g.ColAlignment(4) = 3
    
    g.ColWidth(5) = 700
    g.ColAlignment(5) = 3
    
    g.ColWidth(6) = 1300
    g.ColAlignment(6) = 1
    
    g.ColWidth(7) = 1300
    g.ColAlignment(7) = 3
    
    g.ColWidth(8) = 1000
    g.ColAlignment(8) = 1
    
    g.ColWidth(9) = 3000
    g.ColAlignment(9) = 1
    
    g.ColWidth(10) = 3000
    g.ColAlignment(10) = 1
    
    g.ColWidth(11) = 3000
    g.ColAlignment(11) = 1
    
    g.ColWidth(12) = 8000
    g.ColAlignment(12) = 1

End Sub
Private Sub TxtAte_GotFocus()

    If Len(Trim(TxtAte.Text)) > 0 Then
        TxtAte.SelStart = 0
        TxtAte.SelLength = Len(TxtAte.Text)
    End If
End Sub
Private Sub TxtAte_LostFocus()

    If Not IsDate(TxtAte.Text) Then TxtAte.Text = Format(TxtAte.Text, "00/00/0000")
End Sub
Private Sub TxtDe_GotFocus()

    If Len(Trim(TxtDe.Text)) > 0 Then
        TxtDe.SelStart = 0
        TxtDe.SelLength = Len(TxtDe.Text)
    End If
End Sub
Private Sub TxtDe_LostFocus()

    If Not IsDate(TxtDe.Text) Then TxtDe.Text = Format(TxtDe.Text, "00/00/0000")
End Sub
