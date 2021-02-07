using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

public partial class Manut_Funcionarios : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            CarregaFuncionarios("FUNCIONARIOS_SEL");
            CarregaFuncoes();

            CboSituacao.Items.Add("");
            CboSituacao.Items.Add("Ativo");
            CboSituacao.Items.Add("Inativo");
            LblExportacao.Text = "";
        }
    }

    public void CarregaFuncionarios(string comando)
    {
        //Abre a conexão e executa o comando conforme critérios
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();

        SqlCommand command = new SqlCommand(comando, Conexao);
        SqlDataReader Reader = command.ExecuteReader();
        Grid_Funcionarios_Novo.DataSource = Reader;
        Grid_Funcionarios_Novo.DataBind();
    }

    protected void SelecionaLinhaGrid(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName.Equals("Editar"))
        {
            TxtMatricula.Text = e.CommandArgument.ToString();
            PreencheCampos(TxtMatricula.Text);
        }
    }

    public void PreencheCampos(string Matricula)
    {
        //Abre a conexão e seleciona os dados do registro selecionado
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        SqlCommand command = new SqlCommand("exec FUNCIONARIOS_SEL " + Matricula, Conexao);
        SqlDataReader Reader = command.ExecuteReader();

        if (Reader.HasRows)
        {
            Reader.Read();

            TxtMatricula.Text = Reader["Matricula"].ToString().Trim();
            TxtNome.Text = Reader["Nome"].ToString().Trim();
            CboFuncao.Value = Reader["Funcao"].ToString().Trim();
            TxtCodigoSecao.Text = Reader["Codigo_Secao"].ToString().Trim();
            TxtDescSecao.Text = Reader["Descricao_Secao"].ToString().Trim();
            TxtLocalizacao.Text = Reader["Localizacao"].ToString().Trim();
            TxtEscala.Text = Reader["Horario_Escala_Trabalho"].ToString().Trim();
            TxtDataAdmissao.Text = Reader["Data_Admissao"].ToString().Trim();
            TxtCentroCusto.Text = Reader["Centro_Custo"].ToString().Trim();
            CboSituacao.Value = Reader["Situacao"].ToString().Trim();
        }
    }

    public void CarregaFuncoes()
    {
        //Abre a conexão e seleciona os dados do registro selecionado
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        SqlCommand command = new SqlCommand("select * from FUNC_FUNCIONARIOS", Conexao);
        SqlDataReader Reader = command.ExecuteReader();

        if (Reader.HasRows)
        {
            //Inclua item em branco
            CboFuncao.Items.Add("");
            while (Reader.Read())
            {
                //CboFuncao.Items.Add(Reader["Descricao"].ToString().Trim());
                CboFuncao.Items.Insert(Convert.ToInt32(Reader["Codigo"]), Reader["Descricao"].ToString().Trim());
            }
        }
    }
    protected void CmdSalvar_Click(object sender, EventArgs e)
    {
        string ComandoSQL;

        //Valida os campos da tela
        if (ValidaCampos() == true)
        {
            //Atualiza a tabela de funcionários de acordo com os campos da tela
            SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
            Conexao.Open();

            SqlCommand cmd1 = Conexao.CreateCommand();
            ComandoSQL = " UPDATE Funcionarios SET Nome = '" + TxtNome.Text + "' , ";
            ComandoSQL += " funcao = " + CboFuncao.SelectedIndex + " , ";
            ComandoSQL += " Codigo_Secao = '" + TxtCodigoSecao.Text + "' , ";
            ComandoSQL += " Descricao_Secao = '" + TxtDescSecao.Text + "' , ";
            ComandoSQL += " Localizacao = '" + TxtLocalizacao.Text + "' , ";
            ComandoSQL += " Horario_Escala_Trabalho = '" + TxtEscala.Text + "' , ";
            ComandoSQL += " Data_Admissao = '" + TxtDataAdmissao.Text + "' , ";
            ComandoSQL += " Centro_Custo = '" + TxtCentroCusto.Text + "' , ";
            ComandoSQL += " Situacao = " + CboSituacao.SelectedIndex;
            ComandoSQL += " WHERE Matricula = " + TxtMatricula.Text;

            cmd1.CommandText = ComandoSQL;
            cmd1.ExecuteNonQuery();

            LimpaCampos();
            CarregaFuncionarios("FUNCIONARIOS_SEL");
        }
        else
        {
            //Campos informados não são válidos
            
        }
    }
    protected void CmdLimpar_Click(object sender, EventArgs e)
    {
        LimpaCampos();
    }

    public Boolean ValidaCampos()
    {
        if (TxtMatricula.Text == "")
        {
            return false;
        }
        return true;
    }

    public void LimpaCampos()
    {
        TxtMatricula.Text = "";
        TxtNome.Text = "";
        CboFuncao.Value = "";
        TxtCodigoSecao.Text = "";
        TxtDescSecao.Text = "";
        TxtLocalizacao.Text = "";
        TxtEscala.Text = "";
        TxtDataAdmissao.Text = "";
        TxtCentroCusto.Text = "";
        CboSituacao.SelectedIndex = -1;
    }
    protected void CmdExcluir_Click(object sender, EventArgs e)
    {
        if (TxtMatricula.Text.Length > 0)
        { 
            //Atualizar o status do funcionário para 'Inativo'

        }
    }
    protected void CmdExportar_Click(object sender, EventArgs e)
    {
        LblExportacao.Text = "";

        StreamWriter sw = new StreamWriter(Server.MapPath("a.csv"),true, Encoding.Default);
        StringBuilder Saida = new StringBuilder();

        // Criando o cabeçalho (primeira linha)
        Saida.Append("Matricula" + ";");
        Saida.Append("Nome" + ";");
        Saida.Append("Código Seção" + ";");
        Saida.Append("Descrição Seção" + ";");
        Saida.Append("Localização" + ";");
        Saida.Append("Horário Escala Trabalho" + ";");
        Saida.Append("Data Admissão" + ";");
        Saida.Append("Centro de Custo" + ";");
        Saida.Append("Função" + ";");
        Saida.Append("Situação");
        sw.WriteLine(Saida.ToString());
        Saida.Clear();

        //Abre a conexão e seleciona os dados do registro selecionado
        SqlConnection Conexao = new SqlConnection(BancodeDados.StringConexao);
        Conexao.Open();
        SqlCommand command = new SqlCommand("FUNCIONARIOS_SEL ", Conexao);
        SqlDataReader Reader = command.ExecuteReader();

        if (Reader.HasRows)
        {
            while (Reader.Read())
            {
                for (int Coluna = 0; Coluna < Reader.FieldCount; Coluna++)
                {
                    Saida.Append(Reader[Coluna].ToString().Trim() + ";");
                }
                sw.WriteLine(Saida.ToString().Trim());
                Saida.Clear();
            }
        }
        LblExportacao.Text = "Arquivo Gerado com Sucesso...";
    }
}