using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;

public partial class SIGCT : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            PreencheComboLider();
            CboContrato.Items.Add("");
            CboContrato.Items.Add("CTMARG");
            CboContrato.Items.Add("CTMONSI");
        }
    }

    protected void PreencheComboLider()
    {
        //Abre a conexão
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        string comando;

        comando = "select Nome from funcionarios where lider = '1' order by NOME";

        //Seleciona os nomes dos Lideres e preenche o Combo
        SqlCommand command = new SqlCommand(comando, Conexao);
        SqlDataReader Reader = command.ExecuteReader();
        CboLider.Items.Add("");
        while (Reader.Read())
        {
            CboLider.Items.Add(Reader.GetValue(0).ToString().Trim());
        }
    }
}