using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;

public class BancodeDados
{
    public static string StringConexao = @"Data Source=SP7266ET899;Initial Catalog=DEMANDAS;Integrated Security=True;Connect Timeout=30";
    public SqlConnection Conexao;

    public void ExecutaComandoSQL(string comando)
    {
        //Abre a conexão
        SqlConnection Conexao = new SqlConnection(StringConexao);
        Conexao.Open();

        SqlCommand cmd1 = Conexao.CreateCommand();
        cmd1.CommandText = comando;
        cmd1.ExecuteNonQuery();

        //Fecha a conexão
        Conexao.Close();
    }

    public SqlDataReader SelecionaRegistros(string comando)
    {
        //Abre a conexão
        SqlConnection Conexao = new SqlConnection(StringConexao);
        Conexao.Open();

        //Executa o SELECT na base e retorna o DataReader
        SqlCommand command = new SqlCommand(comando, Conexao);
        SqlDataReader Reader = command.ExecuteReader();
        return Reader;
    }

    public void BulkInsert_Demandas(string path)
    {
        //Abre o arquivo Excel usando conexão OleDB
        OleDbConnection ConexaoExcel = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;");
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = ConexaoExcel;
        if (path.Contains("CTMARG"))
        {
            cmd.CommandText = "Select * from[Todos CTMARG$]";
        }
        else
        {
            cmd.CommandText = "Select * from[Todos CTMONSI$]";
        }

        OleDbDataAdapter objAdapter = new OleDbDataAdapter(cmd);
        ConexaoExcel.Open();
        DbDataReader dr = cmd.ExecuteReader();

        SqlConnection conn = new SqlConnection();
        conn.ConnectionString = StringConexao;
        conn.Open();
        SqlBulkCopy bulkInsert = new SqlBulkCopy(conn);
        bulkInsert.DestinationTableName = "Demandas";
        bulkInsert.WriteToServer(dr);
        ConexaoExcel.Close();
    }

    public void BulkInsert_Servicos(string path)
    {
        //Abre o arquivo Excel usando conexão OleDB -- CTMARG
        OleDbConnection ConexaoExcel = new OleDbConnection("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;");
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = ConexaoExcel;
        cmd.CommandText = "Select * from[CTMARG$]";

        OleDbDataAdapter objAdapter = new OleDbDataAdapter(cmd);
        ConexaoExcel.Open();
        DbDataReader dr = cmd.ExecuteReader();

        SqlConnection conn = new SqlConnection();
        conn.ConnectionString = StringConexao;
        conn.Open();
        SqlBulkCopy bulkInsert = new SqlBulkCopy(conn);
        bulkInsert.DestinationTableName = "Servicos";
        bulkInsert.WriteToServer(dr);
        ConexaoExcel.Close();

        //Abre o arquivo Excel usando conexão OleDB -- CTMONSI
        OleDbCommand cmd2 = new OleDbCommand();
        cmd2.Connection = ConexaoExcel;
        cmd2.CommandText = "Select * from[CTMONSI$]";

        OleDbDataAdapter objAdapter2 = new OleDbDataAdapter(cmd);
        ConexaoExcel.Open();
        DbDataReader dr2 = cmd2.ExecuteReader();

        SqlConnection conn2 = new SqlConnection();
        conn2.ConnectionString = StringConexao;
        conn2.Open();
        SqlBulkCopy bulkInsert2 = new SqlBulkCopy(conn2);
        bulkInsert2.DestinationTableName = "Servicos";
        bulkInsert2.WriteToServer(dr2);
        ConexaoExcel.Close();
    }
}