VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Versao_Procedures 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7164
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   11436
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7164
   ScaleWidth      =   11436
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6780
      Top             =   840
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Versao_Procedures.frx":0000
            Key             =   "TEXT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5100
      Left            =   204
      TabIndex        =   7
      Top             =   1512
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   8996
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome Procedure"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qtd Alterações"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Versão Anterior"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Versão Atual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nome do Banco"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Resetar"
      Height          =   396
      Left            =   120
      TabIndex        =   2
      Top             =   864
      Width           =   1716
   End
   Begin VB.ComboBox cboBancos 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   10956
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de procedures:"
      Height          =   192
      Left            =   4428
      TabIndex        =   6
      Top             =   756
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4428
      TabIndex        =   5
      Top             =   972
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2604
      TabIndex        =   3
      Top             =   972
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resetado dia:"
      Height          =   192
      Left            =   2604
      TabIndex        =   4
      Top             =   756
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco de Dados"
      Height          =   192
      Left            =   108
      TabIndex        =   0
      Top             =   96
      Width           =   1224
   End
End
Attribute VB_Name = "Versao_Procedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_Connection        As New ADODB.Connection


Private Sub cboBancos_Click()

    Dim sStr                As String
    Dim sStrBanco           As String
    Dim rst                 As New ADODB.Recordset
    Dim lListItem           As ListItem
    Dim lColor              As Long
    
    '113 colunas no list
    
    Dim sNome               As String * 40
    Dim sVersaoAnterior     As String * 20
    Dim sVersaoAtual        As String * 20
    Dim sNomeBanco          As String * 10
    Dim sVezesAlteradas     As String * 13
    Dim SCriacao            As String * 12
    
    If cboBancos.ListIndex < 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    sStrBanco = cboBancos.Text


           sStr = "SELECT Max(DataHora) AS DataHora "
    sStr = sStr & "  FROM Reset"
    sStr = sStr & " WHERE IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    
    Set rst = m_Connection.Execute(sStr)
    
    If Not rst.EOF Then
        Label2.Caption = IIf(IsNull(rst!DataHora), "", rst!DataHora)
    End If
    rst.Close

           sStr = "SELECT Count(IdBanco) AS NumProcedures "
    sStr = sStr & "  FROM Procedures "
    sStr = sStr & " WHERE IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    
    Set rst = m_Connection.Execute(sStr)
    
    If Not rst.EOF Then
        Label4.Caption = IIf(IsNull(rst!NumProcedures), "", rst!NumProcedures)
    End If
    rst.Close

    sStr = "SELECT  P.Nome, P.Versao, B.Nome AS NomeBanco, M.Schema_Ver AS VersaoAtual , '0' AS Criacao "
    sStr = sStr & " FROM    DESENV_Versoes..Procedures P, DESENV_Versoes..Banco B, " & cboBancos.Text & "..SysObjects M "
    sStr = sStr & " Where   B.IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    sStr = sStr & " AND     P.IdBanco = B.IdBanco "
    sStr = sStr & " AND     P.Versao <> M.Schema_Ver "
    sStr = sStr & " AND     P.Nome = M.Name "
    sStr = sStr & " AND     M.type = 'P' "
    sStr = sStr & " AND M.status >= 0"
    
    sStr = sStr & " Union All "
    
    sStr = sStr & " SELECT  M.Name, 0 , B.Nome AS NomeBanco, 0 AS VersaoAtual , '1' AS Criacao "
    sStr = sStr & " FROM    DESENV_Versoes..Banco B , " & cboBancos.Text & "..SysObjects M "
    sStr = sStr & " Where B.IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    sStr = sStr & " AND     M.type = 'P' "
    sStr = sStr & " AND M.status >= 0 "
    sStr = sStr & " AND M.NAME not in (select NOME from DESENV_Versoes..procedures where idbanco = " & cboBancos.ItemData(cboBancos.ListIndex) & ")"
    
    sStr = sStr & " Union All "
    
    sStr = sStr & " SELECT  P.Nome, P.Versao, B.Nome AS NomeBanco, 0 AS VersaoAtual , '2' AS Criacao "
    sStr = sStr & " FROM    DESENV_Versoes..Procedures P, DESENV_Versoes..Banco B, " & cboBancos.Text & "..SysObjects M "
    sStr = sStr & " Where B.IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    sStr = sStr & " AND     P.IdBanco = B.IdBanco "
    sStr = sStr & " AND     M.type = 'P' "
    sStr = sStr & " AND M.status > 0 "
    sStr = sStr & " AND     P.Nome *= M.Name "
    sStr = sStr & " AND P.Nome NOT IN (SELECT Name FROM " & cboBancos.Text & "..sysobjects where type = 'P' and Status >= 0) "

    sStr = sStr & " Order By Criacao "

    Set rst = m_Connection.Execute(sStr, 0, adCmdText)
    
    ListView1.ListItems.Clear

    Do While Not rst.EOF()
    
        Set lListItem = ListView1.ListItems.Add(, "Key_" & rst!Nome, rst!Nome, "TEXT")
        
        ListView1.ListItems(rst.AbsolutePosition).SubItems(1) = (rst!VersaoAtual - rst!Versao) / 16
        ListView1.ListItems(rst.AbsolutePosition).SubItems(2) = rst!Versao
        ListView1.ListItems(rst.AbsolutePosition).SubItems(3) = rst!VersaoAtual
        ListView1.ListItems(rst.AbsolutePosition).SubItems(4) = rst!NomeBanco
        ListView1.ListItems(rst.AbsolutePosition).SubItems(5) = IIf(Val(rst!Criacao) = 0, _
                                                                "Alterada", IIf(Val(rst!Criacao) = 1, _
                                                                "Incluida", _
                                                                "Excluida"))

        Select Case Val(rst!Criacao)
        Case 0
            '''''''''''''''''''''''''''''
            'Procedure alterada VERMELHA'
            '''''''''''''''''''''''''''''
            lColor = RGB(255, 0, 0)
        Case 1
            ''''''''''''''''''''''''''
            'Procedure incluida VERDE'
            ''''''''''''''''''''''''''
            lColor = RGB(51, 98, 51)
        Case 2
            '''''''''''''''''''''''''
            'Procedure excluida AZUL'
            '''''''''''''''''''''''''
            lColor = RGB(0, 0, 255)
        End Select
        
        ListView1.ListItems(rst.AbsolutePosition).ForeColor = lColor
        ListView1.ListItems(rst.AbsolutePosition).ListSubItems(1).ForeColor = lColor
        ListView1.ListItems(rst.AbsolutePosition).ListSubItems(2).ForeColor = lColor
        ListView1.ListItems(rst.AbsolutePosition).ListSubItems(3).ForeColor = lColor
        ListView1.ListItems(rst.AbsolutePosition).ListSubItems(4).ForeColor = lColor
        ListView1.ListItems(rst.AbsolutePosition).ListSubItems(5).ForeColor = lColor
        
        rst.MoveNext
    Loop

    rst.Close


    Screen.MousePointer = vbDefault

End Sub


Private Sub cmdReset_Click()

    Dim sStr        As String
    Dim lRetorno    As Long

    
    On Error GoTo Erro_Reset
    
    If cboBancos.ListIndex < 0 Then Exit Sub
    
    
    If MsgBox("Tem certeza ???", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    Label2.Caption = ""
    Label4.Caption = ""
    
           sStr = "DELETE "
    sStr = sStr & "  FROM DESENV_Versoes..Procedures "
    sStr = sStr & " WHERE IdBanco = " & cboBancos.ItemData(cboBancos.ListIndex)
    
    m_Connection.BeginTrans
    
    Call m_Connection.Execute(sStr, lRetorno, adCmdText)
    
    
'    If lRetorno = 0 Then
'        m_Connection.RollbackTrans
'        MsgBox "Não foi possível excluir os registros.", vbExclamation
'        Exit Sub
'    End If
    
           sStr = "INSERT INTO DESENV_Versoes..Procedures "
    sStr = sStr & "SELECT " & cboBancos.ItemData(cboBancos.ListIndex) & ","
    sStr = sStr & "           NAME,"
    sStr = sStr & "           SCHEMA_VER AS VERSAO "
    sStr = sStr & "  FROM " & cboBancos.Text & "..sysobjects"
    sStr = sStr & " WHERE XType = 'P'"
    sStr = sStr & "   AND Category = 0 "
    sStr = sStr & " ORDER by NAME"

    Call m_Connection.Execute(sStr, lRetorno, adCmdText)
    
    If lRetorno = 0 Then
        m_Connection.RollbackTrans
        MsgBox "Não foi possível inserir na base.", vbExclamation
        Exit Sub
    End If
    
           sStr = "INSERT INTO Reset"
    sStr = sStr & "            (IdBanco,DataHora)"
    sStr = sStr & "            VALUES"
    sStr = sStr & "            (" & cboBancos.ItemData(cboBancos.ListIndex) & ","
    sStr = sStr & "            GETDATE())"
    
    Call m_Connection.Execute(sStr, lRetorno, adCmdText)
    
    If lRetorno = 0 Then
        m_Connection.RollbackTrans
        MsgBox "Não foi possível inserir o reset na base.", vbExclamation
        Exit Sub
    End If

    m_Connection.CommitTrans
    
    cboBancos.ListIndex = -1
    
    ListView1.ListItems.Clear

    Exit Sub
Erro_Reset:

    m_Connection.RollbackTrans
    MsgBox Error, vbExclamation


End Sub

Private Sub Form_Load()

    Dim rst         As New ADODB.Recordset
    Dim cmd         As New ADODB.Command
    Dim sServidor   As String
    
    On Error GoTo Erro_Load
    
    sServidor = Command
    If Trim(sServidor) = "" Then
        sServidor = "MDI_NT1"
    End If
    
    
    '''''''''''''''''''''
    'Abertura da conexao'
    '''''''''''''''''''''
    m_Connection.Provider = "SQLOLEDB"
    m_Connection.CursorLocation = adUseClient
    m_Connection.ConnectionString = "SERVER=" & sServidor & ";UID=i;PWD=cau2002;DATABASE=DESENV_VERSOES"
    m_Connection.Open
    
    
    '''''''''''''''''''''''''''''''''''''
    'Pega os bancos de dados adicionados'
    '''''''''''''''''''''''''''''''''''''
    
    Set cmd.ActiveConnection = m_Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GetBancos"
    Set rst = cmd.Execute()
    
    Do While Not rst.EOF()
    
        cboBancos.AddItem rst!Nome
        cboBancos.ItemData(cboBancos.NewIndex) = rst!IdBanco
        rst.MoveNext
    Loop
    
    cboBancos.ListIndex = 0
    rst.Close


    Exit Sub
    
Erro_Load:
    
    MsgBox Error, vbExclamation

End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_Connection.Close
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ListView1.SortKey = ColumnHeader.Index - 1
    
    ListView1.Sorted = True
    


End Sub


