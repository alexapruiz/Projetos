Attribute VB_Name = "basGauge"

Dim TransContador As Integer
Dim EstorContador As Integer

Public Const Transmissao = 0
Public Const Estorno = 1
Public Const Recepcao = 2
Public Const Fininvest = 3

Const LenLabelProgress = 3450

Const CorRecepcao = &H808080        'Cinza
Const CorFininvest = &H808000       'Ciano
Const CorTransmissao = &H800000     'Azul
Const CorEstorno = &HFF&            'Vermelho
Const CorIlegiveis = &H8000&        'Verde
Const CorCorrecao = &H80&           'Vinho
Const CorForaProcess = &H404040     'Cinza
Const CorCSP = &H808000             'Verde piscina
Public Function GaugePos(Optional ByVal Origem As Integer = Transmissao, Optional ByVal pTipoDocto As String = "Desconhecido")
Static Cheques
     
   If Origem = Recepcao Or Origem = Fininvest Then
        TransContador = TransContador + 1
        frmShow.LabelTipoDocto = "Capa"
        frmShow.LabelProcessados = Format(1, "000")
        frmShow.LabelProgressTrans.Width = (LenLabelProgress / Val(frmShow.LabelProgressTrans.Tag)) * TransContador
        frmShow.LabelPercentTrans.Caption = Trim(CInt(((100 / Val(frmShow.LabelProgressTrans.Tag)) * TransContador))) & " % "
        frmShow.LabelTransmitidos = Format(1, "000")
   ElseIf Origem = Estorno Then
        EstorContador = EstorContador + 1
        frmShow.LabelTipoDocto = pTipoDocto
       'frmShow.LabelProcessados = Format(EstorContador, "000")
        frmShow.LabelProcessados = Format(1, "000")
        frmShow.LabelQtde = 1
        frmShow.LabelProgressEstor.Width = (LenLabelProgress / Val(frmShow.LabelProgressEstor.Tag)) * EstorContador
        frmShow.LabelPercentEstor.Caption = Trim(CInt(((100 / Val(frmShow.LabelProgressEstor.Tag)) * EstorContador))) & " % "
       'frmShow.LabelEstornados = Format(Val(frmShow.LabelEstornados.Caption) + 1, "000")
        frmShow.LabelEstornados = Format(1, "000")
   ElseIf Origem = Transmissao Then
        
        If pTipoDocto <> "Cheque Deposito" Then
            Cheques = 0
            TransContador = TransContador + 1
            frmShow.LabelTipoDocto = pTipoDocto
            frmShow.LabelProcessados = Format(TransContador, "000")

            frmShow.LabelProgressTrans.Width = (LenLabelProgress / Val(frmShow.LabelProgressTrans.Tag)) * TransContador
            frmShow.LabelPercentTrans.Caption = Trim(CInt(((100 / Val(frmShow.LabelProgressTrans.Tag)) * TransContador))) & " % "
            frmShow.LabelTransmitidos.Caption = Format(Val(frmShow.LabelTransmitidos.Caption) + 1, "000")
        Else
            Cheques = Cheques + 1
            frmShow.LabelTipoDocto = "Transmitindo: " & pTipoDocto & " = " & Format(Cheques, "000")
        End If
        
   End If
   
   DoEvents
  
End Function
Public Function GaugeInit(ByVal titulo As String, ByVal MaxVal As Integer, Optional ByVal Origem As Integer = Transmissao) As Integer

    DoEvents
    
    If frmShow.cmdFechaCaixa.Caption <> "Aguarde ..." Then
        frmShow.cmdFechaCaixa.Visible = False
    Else
        frmShow.cmdFechaCaixa.Visible = True
    End If
    
    SetColor (Origem)
    
    If titulo <> "" Then
        frmShow.LabelTitulo.Caption = titulo
    End If
       
    If Origem = Transmissao Then
        TransContador = 0
        frmShow.LabelTransmitidos.Caption = "000"
        frmShow.LabelProgressTrans.Tag = MaxVal
        frmShow.LabelProgressTrans.Width = 0
        frmShow.LabelPercentTrans.Caption = " 0 % "
    ElseIf Origem = Estorno Then
        EstorContador = 0
        frmShow.LabelEstornados.Caption = "000"
        frmShow.LabelProgressEstor.Tag = MaxVal
        frmShow.LabelProgressEstor.Width = 0
        frmShow.LabelPercentEstor.Caption = " 0 % "
    ElseIf Origem = Recepcao Or Origem = Fininvest Then
        TransContador = 0
        frmShow.LabelTransmitidos.Caption = "000"
        frmShow.LabelProgressTrans.Tag = MaxVal
        frmShow.LabelProgressTrans.Width = 0
        frmShow.LabelPercentTrans.Caption = " 0 % "
    End If
    
    frmShow.Refresh
End Function
Public Function DestroyGauge()
   frmShow.cmdFechaCaixa.Visible = True
   Call GaugeTitulo(3)
   frmShow.LabelTipoDocto = ""
   frmShow.LabelTransmitidos.Caption = 0
   frmShow.LabelProcessados.Caption = 0
   frmShow.LabelEstornados.Caption = 0
   frmShow.LabelProgressTrans.Width = 0
   frmShow.LabelProgressEstor.Width = 0
   frmShow.LabelPercentTrans.Caption = " 0 % "
   frmShow.LabelPercentEstor.Caption = " 0 % "
   frmShow.LabelQtde.Caption = 0
   
   frmShow.Refresh
End Function
Public Sub GaugeTitulo(pNoTitulo As Integer)
    Select Case pNoTitulo
        Case 0
            frmShow.LabelTitulo.BackColor = CorTransmissao
        Case 1
            frmShow.LabelTitulo.BackColor = CorCorrecao
            frmShow.LabelTitulo.Caption = "Enviando Capa para Correção AG/CC"
        Case 2
            frmShow.LabelTitulo.BackColor = CorIlegiveis
            frmShow.LabelTitulo.Caption = "Enviando Capa para Ilegíveis"
        Case 3
            frmShow.LabelTitulo.BackColor = CorForaProcess
            frmShow.LabelTitulo.Caption = "Fora de Processamento"
        Case 4
            frmShow.LabelTitulo.BackColor = CorCSP
            frmShow.LabelTitulo.Caption = "Enviando Capa para CSP"
    End Select

End Sub
Sub SetColor(pOrigem)
    Call GaugeTitulo(0)

    If pOrigem = Transmissao Then
        frmShow.LabelInstrucao.BackColor = CorTransmissao
        frmShow.LabelProgressTrans.BackColor = CorTransmissao
    ElseIf pOrigem = Recepcao Then
        frmShow.LabelInstrucao.BackColor = CorRecepcao
        frmShow.LabelProgressTrans.BackColor = CorRecepcao
    ElseIf pOrigem = Fininvest Then
        frmShow.LabelInstrucao.BackColor = CorFininvest
        frmShow.LabelProgressTrans.BackColor = CorFininvest
    ElseIf pOrigem = Estorno Then
        frmShow.LabelInstrucao.BackColor = CorEstorno
        frmShow.LabelProgressEstor.BackColor = CorEstorno
    End If
    
End Sub
