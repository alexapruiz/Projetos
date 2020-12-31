using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

public partial class Consumo_USTs : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Opt_Unidade.SelectedIndex = -1;
            CmdGrids.Visible = false;
            Titulo_Grid1.Text = "";
            Titulo_Grid2.Text = "";
            PreencheComboPeriodo();
        }
    }

    protected void PreencheComboPeriodo()
    {
        DateTime Data;
        DateTime DataInicial;
        DateTime Hoje = DateTime.Now;

        DataInicial = Convert.ToDateTime("01/02/2014");
        Data = Convert.ToDateTime(Hoje);

        //Preenche o combo com o periodo atual e anteriores até 02/2014
        while (Data > DataInicial)
        {
            CboPeriodo.Items.Add(Convert.ToString(Convert.ToString(Data).Substring(3, 7)));
            Data = Data.AddMonths(-1);
        }
    }

    protected void CmdPesquisar_Click(object sender, EventArgs e)
    {
        PesquisaDemandas("Grid1");
    }

    protected void PesquisaDemandas(string Grid)
    {
        string filtro_padrao = " WHERE STATUS <> 'Cancelado' ";
        string filtro_unidade = "";
        string filtro_periodo = "";
        string Unidade_Selec = "";

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
            filtro_unidade = "AND UNIDADE IN ('" + filtro_unidade.Trim() + ")";
        }

        //Verifica o Tipo de Visão
        if (CboTipoVisao.Value == "Previsto")
        {
            filtro_periodo = " AND periodo_prev = '" + Convert.ToDateTime(CboPeriodo.Value).ToString("yyyy/MM") + "' ";
        }
        else
        {
            filtro_periodo = " AND periodo_real = '" + Convert.ToDateTime(CboPeriodo.Value).ToString("yyyy/MM") + "' ";
        }

        //Abre a conexão
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        string comando;
        string comando_agrup;

        comando = "select Contrato , Segmento , sum(total_ust) as 'Qtde USTs' ";
        comando += " from Demandas ";

        comando_agrup = " group by CONTRATO , Segmento  ";
        comando_agrup += " order by CONTRATO asc , Segmento  ";

        //Executa o SELECT na base e preenche o grid
        SqlCommand command = new SqlCommand(comando + filtro_padrao + filtro_periodo + filtro_unidade + comando_agrup, Conexao);
        SqlDataReader Reader = command.ExecuteReader();

        if (Grid == "Grid1")
        {
            Grid_Demandas1.DataSource = Reader;
            Grid_Demandas1.DataBind();

            for (int i = 0; i <= 2; i++)
            {
                if (Opt_Unidade.Items[i].Selected == true)
                {
                    Unidade_Selec = Unidade_Selec + Opt_Unidade.Items[i].Value.Trim() + " , ";
                }
            }
            //Retirando o último caracter ','
            Unidade_Selec = Unidade_Selec.ToString().Trim().Substring(0, Unidade_Selec.Length - 2);

            Titulo_Grid1.Text = "Unidade: " + Unidade_Selec + " - Tipo: " + CboTipoVisao.Value + " - Período: " + CboPeriodo.Value;
            CmdGrids.Visible = true;
        }
        else
        {
            Grid_Demandas2.DataSource = Reader;
            Grid_Demandas2.DataBind();
            Titulo_Grid2.Text = Titulo_Grid1.Text;
        }
    }
    protected void CmdGrids_Click(object sender, EventArgs e)
    {
        Grid_Demandas1.DataSource = "";
        Grid_Demandas1.DataBind();
        PesquisaDemandas("Grid2");
        Titulo_Grid1.Text = "";
    }
    protected void CmdLimparPesquisa_Click(object sender, EventArgs e)
    {
        Grid_Demandas1.DataSource = "";
        Grid_Demandas1.DataBind();
        Grid_Demandas2.DataSource = "";
        Grid_Demandas2.DataBind();
        Titulo_Grid1.Text = "";
        Titulo_Grid2.Text = "";
        Opt_Unidade.SelectedIndex = -1;
        CboPeriodo.SelectedIndex = 0;
        CmdGrids.Visible = false;
    }
}