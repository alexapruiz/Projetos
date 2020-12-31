using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;

public partial class Manutencao : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Label1.Text = "Selecione o arquivo...";
    }

    protected void CmdImportar_Click(object sender, EventArgs e)
    {
        ImportarRegistros();

        //if (CmdImportar.Text == "Importar")
        //{
        //    //Session.Add("Parado", "0");
        //    CmdImportar.Text = "Parar";
        //    System.Threading.Thread Th = new System.Threading.Thread(ImportarRegistros);
        //    Th.Start();
        //}
        //else
        //{
        //    CmdImportar.Text = "Importar";
        //    Session.Add("Parado", "1");
        //    pnlStatus.Width = Unit.Pixel(0);
        //}
    }
    protected void CmdZerar_Click(object sender, EventArgs e)
    {
        if (CboTipoImportacao.SelectedIndex == 0)
        {
            //Limpa a tabela DEMANDAS
            BancodeDados BD = new BancodeDados();
            BD.ExecutaComandoSQL("TRUNCATE TABLE DEMANDAS");
        }
        else
        {
            //Limpa a tabela DEMANDAS
            BancodeDados BD = new BancodeDados();
            BD.ExecutaComandoSQL("TRUNCATE TABLE SERVICOS");
        }
    }

    private void IniciarProcessoLongo()
    {
        for (long Cont = 0; Cont <= 100; Cont++)
        {
            Session.Add("Status", Cont);
            for (int Cont2 = 0; Cont2 <= 100000000; Cont2++)
            {
                if (CmdImportar.Text == "Importar")
                {
                    return;
                }

                if (Session["Parado"].ToString() == "1")
                {
                    return;
                }
                //pnlStatus.Width = Unit.Pixel(Convert.ToInt32(Session["Status"]) * 5);
            }
        }
    }

    public void ImportarRegistros()
    {
        string TipoDemanda;
        string Agora;
        string ComandoSQL;
        //decimal Contador = 0;
        //decimal Acumulador = 0;

        if (FileUpload1.HasFile)
        {
            //Hoje = DateTime.Now.Replace("/", "_");
            Agora = DateTime.Now.ToString();
            Agora = Agora.Replace("/", "_");
            Agora = Agora.Replace(":", "_");
            Agora = Agora.Replace(" ", "_");

            //Cria cópia temporária do arquivo selecionado
            string path = string.Concat((Server.MapPath("~/temp/" + FileUpload1.FileName)));
            FileUpload1.PostedFile.SaveAs(path);

            BancodeDados BD = new BancodeDados();

            //Verificar se a importação é para Demandas ou Serviços
            if (CboTipoImportacao.SelectedIndex == 0)
            {
                //Importaçao de demandas
                //Realiza backup da tabela atual
                //BD.ExecutaComandoSQL("SELECT * INTO DEMANDAS_" + Agora + " FROM DEMANDAS");

                OleDbConnection ConexaoExcel = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES';");

                if (FileUpload1.FileName.Contains("CTMARG"))
                {
                    TipoDemanda = "CTMARG";
                }
                else
                {
                    TipoDemanda = "CTMONSI";
                }

                try
                {
                    Label1.Text = "0%";

                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [" + TipoDemanda + "$]", ConexaoExcel);
                    DataSet ds = new DataSet();
                    ConexaoExcel.Open();
                    adapter.Fill(ds);

                    //Contador = (Convert.ToDecimal(100.000) / Convert.ToDecimal(ds.Tables[0].Rows.Count));
                    foreach (DataRow linha in ds.Tables[0].Rows)
                    {
                        ComandoSQL = "INSERT INTO DEMANDAS (ID,Resumo,Status,Tags,Quantidade,Rejeitados,Complexidade,Unidade,Equipe,Atividade,Servico,Data_Criacao,Prazo_Final,Data_Resolucao,CONTRATO) VALUES ('";
                        ComandoSQL += linha["ID"].ToString() + "','";
                        ComandoSQL += linha["Resumo"].ToString().Replace("'", "").Replace(",", ";") + "','";
                        ComandoSQL += linha["Status"].ToString() + "','";
                        ComandoSQL += linha["Tags"].ToString().Replace(",", ";") + "','";
                        ComandoSQL += linha["Quantidade"].ToString() + "','";
                        ComandoSQL += linha["Rejeitados"].ToString() + "','";
                        ComandoSQL += linha["Complexidade"].ToString() + "','";

                        if (TipoDemanda == "CTMARG")
                        {
                            ComandoSQL += linha["Unidade Solicitante em Métodos e Processos"].ToString() + "','";
                            ComandoSQL += linha["Segmento Solicitante em Métodos e Processos"].ToString() + "','";
                            ComandoSQL += linha["Atividades em Método e Processo"].ToString() + "','";
                            ComandoSQL += linha["Serviços Métodos e Processos"].ToString() + "','";
                        }
                        else
                        {
                            ComandoSQL += linha["Unidade Solicitante Especializada"].ToString() + "','";
                            ComandoSQL += linha["Segmento Solicitante Especializado"].ToString() + "','";
                            ComandoSQL += linha["Atividade Especializada"].ToString() + "','";
                            ComandoSQL += linha["Serviço Especializado"].ToString() + "','";
                        }

                        if (linha["Data de Criação"].ToString().Length > 0)
                        {
                            ComandoSQL += Convert.ToDateTime(linha["Data de Criação"].ToString()) + "','";
                        }
                        else
                        {
                            ComandoSQL += "','";
                        }

                        if (linha["Prazo Final"].ToString().Length > 0)
                        {
                            ComandoSQL += Convert.ToDateTime(linha["Prazo Final"].ToString()) + "','";
                        }
                        else
                        {
                            ComandoSQL += "','";
                        }

                        if (linha["Data de Resolução"].ToString().Length > 0)
                        {
                            ComandoSQL += Convert.ToDateTime(linha["Data de Resolução"].ToString()) + "','";
                        }
                        else
                        {
                            ComandoSQL += "','";
                        }

                        ComandoSQL += TipoDemanda + "')";

                        BD.ExecutaComandoSQL(ComandoSQL);

                        //Acumulador = Acumulador + Contador;
                        //Session.Add("Status", Acumulador);
                        //pnlStatus.Width = Unit.Pixel(Convert.ToInt32(Acumulador));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erro ao acessar os dados: " + ex.Message);
                }
                finally
                {
                    ConexaoExcel.Close();
                }

                //Exclui arquivo temporário
                Array.ForEach(Directory.GetFiles((Server.MapPath("~/temp/"))), File.Delete);

                //Corrige Campo Quantidade
                BD.ExecutaComandoSQL("UPDATE Demandas SET Quantidade = 1 WHERE Quantidade IS NULL or Quantidade = 0");

                //Preenche o campo PERFIL
                BD.ExecutaComandoSQL("update Demandas set PERFIL = 1 where Complexidade = 'Alta'");
                BD.ExecutaComandoSQL("update Demandas set PERFIL = 2 where Complexidade <> 'Alta'");

                //Preenche o Campo PERIODO_PREV
                BD.ExecutaComandoSQL("EXEC ATUALIZA_PERIODO_PREV_DEMANDAS");

                //Preenche o Campo PERIODO_REAL
                BD.ExecutaComandoSQL("EXEC ATUALIZA_PERIODO_REAL_DEMANDAS");

                //Altera a complexidade de 'N/A' para 'Baixa'
                BD.ExecutaComandoSQL("update demandas set Complexidade = 'Baixa' where Complexidade = 'N/A'");

                //Calcula as UST's de cada demanda
                BD.ExecutaComandoSQL("EXEC ATUALIZA_SERVICOS_DEMANDAS " + TipoDemanda);

                //Define todas as demandas como 'Sem Classificação'
                BD.ExecutaComandoSQL("update demandas set SEGMENTO = 'Sem Classificação'");

                //Preenche o campo Segmento
                BD.ExecutaComandoSQL("EXEC ATUALIZA_SEGMENTOS_DEMANDAS");

                //Preenche o campo TOTAL_UST
                BD.ExecutaComandoSQL("update Demandas set TOTAL_UST = Quantidade * UST");

                Label1.Text = "Importação realizada com sucesso...";
            }
            else if (CboTipoImportacao.SelectedIndex == 1)
            {
                // Importação de Serviços
                OleDbConnection ConexaoExcel = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES';");

                if (FileUpload1.FileName.Contains("CTMARG"))
                {
                    TipoDemanda = "CTMARG";
                }
                else
                {
                    TipoDemanda = "CTMONSI";
                }

                try
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [" + TipoDemanda + "$]", ConexaoExcel);
                    DataSet ds = new DataSet();
                    ConexaoExcel.Open();
                    adapter.Fill(ds);

                    foreach (DataRow linha in ds.Tables[0].Rows)
                    {
                        ComandoSQL = "INSERT INTO SERVICOS (NUMERO,CONTRATO,ATIVIDADE,SERVICO,ESFORCO_BAIXA,ESFORCO_MEDIA,ESFORCO_ALTA) VALUES ('";
                        ComandoSQL += linha["Numero"].ToString() + "','";
                        ComandoSQL += linha["Contrato"].ToString() + "','";
                        ComandoSQL += linha["Atividade"].ToString().Replace("'", "") + "','";

                        if (linha["Serviço"].ToString().Length > 300)
                        {
                            ComandoSQL += linha["Serviço"].ToString().Replace("'", "").Substring(0, 299) + "','";
                        }
                        else
                        {
                            ComandoSQL += linha["Serviço"].ToString().Replace("'", "") + "','";
                        }

                        ComandoSQL += linha["Duração Baixa"].ToString() + "','";
                        ComandoSQL += linha["Duração Média"].ToString() + "','";
                        ComandoSQL += linha["Duração Alta"].ToString() + "')";

                        BD.ExecutaComandoSQL(ComandoSQL);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erro ao acessar os dados: " + ex.Message);
                }
                finally
                {
                    ConexaoExcel.Close();
                }
            }
        }
        else
        {
            Label1.Text = "Um arquivo deve ser selecionado...";
        }
    }
}