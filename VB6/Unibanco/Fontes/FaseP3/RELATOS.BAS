Attribute VB_Name = "RELATORIOS"
Dim linha As Integer, pag As Integer, rep As Integer

Dim valor_caption As String
Dim str_formatada As String * 13

Dim val_env As Double


Sub Cabec1()
   
   Dim diamov As String

   pag = pag + 1
   If pag > 1 Then
      Printer.NewPage
   End If
   diamov = Mid$(data_inverte, 5, 2) + " / " + Mid$(data_inverte, 3, 2) + " / " + Mid$(data_inverte, 1, 2)
   
   Printer.Print " PROSERVVI                TOTAIS DE ENVELOPES RECEBIDOS            " + "  Pagina:" + Str$(pag)
   Printer.Print
   Printer.Print "                     Movimento de: " + diamov + "            " + Format$(Now, "dd-mm-yyyy") + "  " + Time$
   Printer.Print
   Printer.Print String(80, "-")
   Printer.Print "      Agencia n°              Qtde. Envelopes                  Valor total - R$"
   Printer.Print String(80, "-")
   Printer.Print
   linha = 9

End Sub

Sub cabec2()
   
   Dim diamov As String

   pag = pag + 1
   If pag > 1 Then
      Printer.NewPage
   End If
   diamov = Mid$(data_inverte, 5, 2) + " / " + Mid$(data_inverte, 3, 2) + " / " + Mid$(data_inverte, 1, 2)
   
   Printer.Print " PROSERVVI                    TRANSAÇÕES DO DIA                  " + "  Pagina:" + Str$(pag)
   Printer.Print
   Printer.Print "                     Movimento de: " + diamov + "            " + Format$(Now, "dd-mm-yyyy") + "  " + Time$
   Printer.Print
   Printer.Print String(80, "-")
   Printer.Print "    Código da Transação             Qtde.                  Valor total - R$"
   Printer.Print String(80, "-")
   Printer.Print
   linha = 9

End Sub


Sub cabec3()
   
   Dim diamov As String

   pag = pag + 1
   If pag > 1 Then
      Printer.NewPage
   End If
   diamov = Mid$(data_inverte, 5, 2) + " / " + Mid$(data_inverte, 3, 2) + " / " + Mid$(data_inverte, 1, 2)
   
   Printer.Print " PROSERVVI                 TRANSAÇÕES DO DIA - POR AGENCIA          " + "  Pagina:" + Str$(pag)
   Printer.Print
   Printer.Print "                     Movimento de: " + diamov + "            " + Format$(Now, "dd-mm-yyyy") + "  " + Time$
   Printer.Print
   Printer.Print "                                    Agencia: " + rel_ag_orig
   Printer.Print
   Printer.Print String(80, "-")
   Printer.Print "    Código da Transação             Qtde.                  Valor total - R$"
   Printer.Print String(80, "-")
   Printer.Print
   linha = 11

End Sub


Sub cabec4()
   
   Dim diamov As String

   pag = pag + 1
   If pag > 1 Then
      Printer.NewPage
   End If
   diamov = Mid$(data_inverte, 5, 2) + " / " + Mid$(data_inverte, 3, 2) + " / " + Mid$(data_inverte, 1, 2)
   
   Printer.Print " PROSERVVI                  MALOTE EMPRESA - POR AGENCIA             " + "  Pagina:" + Str$(pag)
   Printer.Print
   Printer.Print "                     Movimento de: " + diamov + "            " + Format$(Now, "dd-mm-yyyy") + "  " + Time$
   Printer.Print
   Printer.Print "                                 Agencia: " + rel_ag_orig
   Printer.Print
   Printer.Print String(80, "-")
   Printer.Print " Número da Conta   Qtde. Doctos            Número da Conta   Qtde. Doctos"
   Printer.Print String(80, "-")
   Printer.Print
   linha = 11

End Sub


Sub Formata_Valor()
   
   ' Carrega imagem e valor formatado na tela
   valor_caption = Format$((val_env / 100), "###,###,###.00")
   
   str_formatada = ""     'limpa string para guardar valor formatado
   rep = 1                'contador de caracteres a serem formatados
   
   Do
      If (Mid$(valor_caption, rep, 1) = "") Then   'verifica término da string
         Exit Do
      End If
      rep = rep + 1
   Loop While (rep < 14)   'tamanho máximo da string a ser formatada
   rep = rep - 1
   
   ' formata à direita
   Mid$(str_formatada, 13 - rep + 1, rep) = Mid$(valor_caption, 1, rep)
   valor_caption = str_formatada    'atualiza valor_caption com dado formatado a direita

End Sub

Sub rel1()
   
   Dim Ag_rel As String
   Dim qtd_env As Integer
   
   Screen.MousePointer = 11   'ampulheta
   
   pag = 0
   linha = 60
   qtd_env = 0
   val_env = 0
   
   rp_dae.DtAgEnv = data_inverte
   lixo = GTEREC(RecupDtAgEnv, rp_dae, reg_recup)
   If (lixo <> 0) Or (reg_recup.Data_Movimento <> data_inverte) Then
      Screen.MousePointer = 0         'default
      Beep
      MsgBox "Não existe nenhum Envelope cadastrado nesta Data !", 16, "Atenção"
      lixo = CLISAM()
      Exit Sub
   End If
   
   Ag_rel = reg_recup.Ag_orig
   
   Do
      If linha > 58 Then
          Cabec1
      End If
      
      If (Ag_rel <> reg_recup.Ag_orig) Then
      
         Formata_Valor
         Printer.Print Space(8) + Ag_rel + Space(25) + Format$(qtd_env, "#,###") + Space(25) + valor_caption
         Printer.Print
         linha = linha + 2
         qtd_env = 0
         val_env = 0
         Ag_rel = reg_recup.Ag_orig
      End If
      
      qtd_env = qtd_env + 1
      val_env = val_env + Val(reg_recup.Vl_Conta)
      
      lixo = NXTREC(RecupDtAgEnv, reg_recup)
      
      If (lixo <> 0) Or (reg_recup.Data_Movimento <> data_inverte) Then
         Formata_Valor
         Printer.Print Space(8) + Ag_rel + Space(25) + Format$(qtd_env, "#,###") + Space(25) + valor_caption
         Printer.Print
         Exit Do
      End If
   Loop
   
   Printer.Print
   Printer.Print String(80, "-")
   Printer.EndDoc
        
   Screen.MousePointer = 0

End Sub

Sub rel3()
   
   Dim tran_dia As String
   Dim qtd_tr As Integer
   
   Screen.MousePointer = 11   'ampulheta
   
   pag = 0
   linha = 60
   qtd_tr = 0
   val_env = 0
   
   Et_dt.DtTr = data_inverte
   lixo = GTEREC(EstDTr, Et_dt, reg_est)
   If (lixo <> 0) Or (reg_est.Data_Movimento <> data_inverte) Then
      Screen.MousePointer = 0         'default
      Beep
      MsgBox "Não existe nenhum Movimento cadastrado nesta Data !", 16, "Atenção"
      lixo = CLISAM()
      Exit Sub
   End If
   
   tran_dia = reg_est.Trans
   
   Do
      If linha > 58 Then
          cabec2
      End If
      
      If (tran_dia <> reg_est.Trans) Then
      
         Formata_Valor
         Printer.Print Space(8) + tran_dia + Space(25) + Format$(qtd_tr, "#,###") + Space(25) + valor_caption
         Printer.Print
         linha = linha + 2
         qtd_tr = 0
         val_env = 0
         tran_dia = reg_est.Trans
      End If
      
      qtd_tr = qtd_tr + Val(reg_est.Qtde)
      val_env = val_env + Val(reg_est.valor)
      
      lixo = NXTREC(EstDTr, reg_est)
      
      If (lixo <> 0) Or (reg_est.Data_Movimento <> data_inverte) Then
         Formata_Valor
         Printer.Print Space(8) + tran_dia + Space(25) + Format$(qtd_tr, "#,###") + Space(25) + valor_caption
         Printer.Print
         Exit Do
      End If
   Loop
   
   Printer.Print
   Printer.Print String(80, "-")
   Printer.EndDoc
        
   Screen.MousePointer = 0


End Sub

Sub rel2()

   Screen.MousePointer = 11   'ampulheta
   
   pag = 0
   linha = 60
   val_env = 0
   
   Et_dat.DtAgTr = data_inverte + rel_ag_orig
   lixo = GTEREC(EstDAgTr, Et_dat, reg_est)
   If (lixo <> 0) Or (reg_est.Data_Movimento <> data_inverte) Then
      Screen.MousePointer = 0         'default
      Beep
      MsgBox "Não existe nenhum Movimento cadastrado nesta Data !", 16, "Atenção"
      lixo = CLISAM()
      Exit Sub
   End If
   
   Do
      If linha > 58 Then
          cabec3
      End If
      
      qtd_tr = Val(reg_est.Qtde)
      val_env = Val(reg_est.valor)
      Formata_Valor
      Printer.Print Space(8) + reg_est.Trans + Space(25) + Format$(qtd_tr, "#,###") + Space(25) + valor_caption
      Printer.Print
      linha = linha + 2
      tran_dia = reg_est.Trans
      
      lixo = NXTREC(EstDAgTr, reg_est)
      
      If (lixo <> 0) Or (reg_est.Data_Movimento <> data_inverte) Or (reg_est.agencia <> rel_ag_orig) Then
         Exit Do
      End If
   
   Loop
   
   Printer.Print
   Printer.Print String(80, "-")
   Printer.EndDoc
        
   Screen.MousePointer = 0

End Sub


Sub rel4()
   
   Screen.MousePointer = 11   'ampulheta
   
   Dim linha_rel As String
   Dim coluna As Integer
   
   pag = 0
   linha = 60
     
     
   rp_dae.DtAgEnv = data_inverte + rel_ag_orig
   
   lixo = GTEREC(RecupDtAgEnv, rp_dae, reg_recup)
   If (lixo <> 0) Or (reg_recup.Data_Movimento <> data_inverte) Then
      Screen.MousePointer = 0         'default
      Beep
      MsgBox "Não existe nenhum Envelope cadastrado nesta Data !", 16, "Atenção"
      lixo = CLISAM()
      Exit Sub
   End If
   
   linha_rel = Space(3)
   coluna = 0
   
   Do
      If linha > 58 Then
          cabec4
      End If
      
      If Val(Cta_orig) <> 0 Then
         
         linha_rel = linha_rel + reg_recup.Cta_orig + Space(15) + reg_recup.Qtd_docs + Space(17)
         coluna = coluna + 1
         
         If coluna = 2 Then
            Printer.Print linha_rel
            linha = linha + 1
            coluna = 0
            linha_rel = Space(3)
         End If
      
      End If
      
      lixo = NXTREC(RecupDtAgEnv, reg_recup)
      If (lixo <> 0) Or (reg_recup.Data_Movimento <> data_inverte) Then
            Exit Do        'leitura da próxima caixa
      End If
   Loop
   
   Printer.Print linha_rel
   Printer.EndDoc
        
   Screen.MousePointer = 0

End Sub


Sub Relacao_Baixas()
   
   Dim linha_rel As String
   Dim coluna As Integer
   
   Screen.MousePointer = 11   'ampulheta
   
   pag = 0
   linha = 60
   coluna = 0
   
   pg_dtp.Dtp = data_inverte
   pg_dtp.Deslo = 0
   
   lixo = GTEREC(ProgDtp, pg_dtp, reg_prog)
   If (lixo <> 0) Or (reg_prog.Data_Programada <> data_inverte) Then
      Screen.MousePointer = 0
      Beep
      MsgBox "Não existe nenhum Envelope cadastrado nesta Data !", 16, "Atenção"
      lixo = CLISAM()
      Exit Sub
   End If
   
   Do While (lixo = 0) And (reg_prog.Data_Programada = data_inverte)
      
      If linha > 58 Then
         If pag > 0 Then
            Printer.NewPage
         End If
    
         pag = pag + 1
         Printer.Print
         Printer.Print "PROSERVVI - Cash Image          Baixa Programados" + Space(10) + " Pagina: " + Str$(pag)
         Printer.Print
         Printer.Print Space(5) + "Relação dos Envelopes Programados para: " + Mid$(data_inverte, 5, 2) + "/" + Mid$(data_inverte, 3, 2) + "/" + Mid$(data_inverte, 1, 2) + Space(13) + Time$
         Printer.Print String(86, "-")
         linha = 5
         coluna = 0
         
      End If
      
      linha_rel = linha_rel + reg_prog.Nro_Envelope + Space(5)
      coluna = coluna + 1
         
      If coluna = 4 Then
         Printer.Print linha_rel
         linha = linha + 1
         coluna = 0
         linha_rel = Space(3)
      End If
         
      lixo = NXTREC(ProgDtp, reg_prog)
   Loop
   
   Printer.Print linha_rel
   Printer.EndDoc
        
   Screen.MousePointer = 0

End Sub

