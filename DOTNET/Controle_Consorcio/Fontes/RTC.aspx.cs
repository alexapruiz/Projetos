using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Net;

public partial class RTC : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Opt_Unidade.SelectedIndex = -1;
            LblResumoConsulta.Text = "";
            PreencheCombosPeriodo();
        }
    }

    protected void CmdPesquisar_Click(object sender, EventArgs e)
    {
        //Consulta demandas conforme parâmetros informados pelo usuário
        ConsultaDemandas("");
    }

    protected void PreencheCombosPeriodo()
    {
        DateTime Data;
        DateTime DataInicial;
        DateTime Hoje = DateTime.Now;

        DataInicial = Convert.ToDateTime("01/02/2014");
        Data = Convert.ToDateTime(Hoje);

        //Preenche o combo com o periodo atual e anteriores até 02/2014
        while (Data > DataInicial)
        {
            CboPeriodoPrevisto.Items.Add(Convert.ToString(Convert.ToString(Data).Substring(3, 7)));
            CboPeriodoReal.Items.Add(Convert.ToString(Convert.ToString(Data).Substring(3, 7)));
            Data = Data.AddMonths(-1);
        }
    }

    public void ConsultaDemandas(string Ordenacao)
    {
        string comando = "";
        string filtro_periodo = "";
        DateTime DataValida;

        if (CboTipoConsulta.SelectedIndex == 0)
        {
            //Pesquisar demandas sem tag
            comando = "Select CONTRATO , ID , Equipe , Status, Resumo , substring(periodo_prev,6,2) + '/' + substring(periodo_prev,1,4) as 'Período Previsto' , substring(periodo_real,6,2) + '/' + substring(periodo_real,1,4) as 'Período Real' from DEMANDAS where (TAGS is null or tags = '') and Status not in ('Cancelado')";
        }
        else if (CboTipoConsulta.SelectedIndex == 1)
        {
            //Pesquisar demandas atrasadas
            comando = "Select CONTRATO , ID , RESUMO , STATUS , Quantidade , TOTAL_UST , DATA_CRIACAO , PRAZO_FINAL , substring(periodo_prev,6,2) + '/' + substring(periodo_prev,1,4) as 'Período Previsto' from DEMANDAS where Prazo_Final < '" + DateTime.Now.ToString("dd/MM/yyyy");
            comando += "' and status not in ('Validado','Entregue','Cancelado', 'Suspenso','Aceito','Recebido','Pendente')";
        }
        else if (CboTipoConsulta.SelectedIndex == 2)
        {
            //Pesquisar demandas com periodo previsto diferente do periodo real
            comando = "Select CONTRATO , ID , Equipe , Status, Resumo , substring(periodo_prev,6,2) + '/' + substring(periodo_prev,1,4) as 'Período Previsto' , substring(periodo_real,6,2) + '/' + substring(periodo_real,1,4) as 'Período Real' from DEMANDAS where Status not in ('Cancelado') and periodo_prev <> periodo_real ";
        }

        string filtro_unidade = "";
        string filtro_periodo1 = "";
        string filtro_periodo2 = "";
        string filtro_contrato = "";

        //Verifica se foi informada uma unidade
        if (Opt_Unidade.SelectedIndex != -1)
        {
            for (int i = 0; i <= 2; i++)
            {
                if (Opt_Unidade.Items[i].Selected == true)
                {
                    filtro_unidade = filtro_unidade + "'" + Opt_Unidade.Items[i].Value.Trim().Replace(" / ", "") + "',";
                }
            }
            //Retirando o último caracter ','
            filtro_unidade = filtro_unidade.Substring(1, filtro_unidade.Length - 2);
            filtro_unidade = "AND UNIDADE IN ('" + filtro_unidade.Trim() + ") ";
        }

        //Verifica se foi informado um contrato
        if (CboContrato.SelectedIndex > 0)
        {
            filtro_contrato = " AND CONTRATO = '" + CboContrato.Value + "' ";
        }

        //Verifica se foi informado um periodo (Real ou previsto)
        if (CboPeriodoPrevisto.SelectedIndex > 0)
        {
            filtro_periodo1 = filtro_periodo1 + " AND periodo_prev = '" + Convert.ToDateTime(CboPeriodoPrevisto.Value).ToString("yyyy/MM") + "' ";
        }

        if (CboPeriodoReal.SelectedIndex > 0)
        {
            filtro_periodo2 = filtro_periodo2 + " AND periodo_real = '" + Convert.ToDateTime(CboPeriodoReal.Value).ToString("yyyy/MM") + "' ";
        }

        //Verifica se foi informado um periodo
        if (TxtPeriodo.Value != "")
        {
            string Data1 = "";
            string Data2 = "";

            //Separando as datas
            try
            {
                Data1 = TxtPeriodo.Value.Substring(0, 10);
                Data2 = TxtPeriodo.Value.Substring(13, 10);

                if (DateTime.TryParse(Data1, out DataValida) || DateTime.TryParse(Data2, out DataValida))
                {
                    filtro_periodo = filtro_periodo + " AND prazo_final BETWEEN '" + Data1 + " 00:00:00' AND '" + Data2 + " 23:59:59'";
                }
                else
                {
                    TxtPeriodo.Value = "";
                }
            }
            catch
            {
                Data1 = "";
                Data2 = "";
                filtro_periodo = "";
                TxtPeriodo.Value = "";
            }
        }

        //Monta o comando completo
        string Comando_Completo = "";
        if (Ordenacao != "")
        {
            //Verifica se o usuário solicitou ordenação por campos de Data, onde as colunas do grid possuem nomes diferentes da tabela
            if (Ordenacao == "Período Previsto")
            {
                Ordenacao="periodo_prev";
            }

            if (Ordenacao == "Período Real")
            {
                Ordenacao = "periodo_real";
            }

            Comando_Completo = comando + filtro_periodo1 + filtro_periodo2 + filtro_unidade + filtro_contrato + filtro_periodo + " ORDER BY " + Ordenacao;
        }
        else
        {
            Comando_Completo = comando + filtro_periodo1 + filtro_periodo2 + filtro_unidade + filtro_contrato + filtro_periodo + " ORDER BY CONTRATO , ID , RESUMO , STATUS , PRAZO_FINAL";
        }

        //Abre a conexão e executa o comando conforme critérios
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        SqlCommand command = new SqlCommand(Comando_Completo, Conexao);
        SqlDataReader Reader = command.ExecuteReader();
        Grid_Demandas.DataSource = Reader;
        Grid_Demandas.DataBind();
        LblResumoConsulta.Text = "Foram encontrados " + Grid_Demandas.Rows.Count.ToString() + " registros!!!";       
    }

    protected void OrdenaGridDemandas(object sender, GridViewSortEventArgs e)
    {
        ConsultaDemandas(e.SortExpression);
    }
}