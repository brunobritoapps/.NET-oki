using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

public partial class Adm_rptcoletasaprovadas : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String sfiltro = Request.QueryString["filtra"];

        this.ExportToExcel(sfiltro);

    }

    /// <summary>
    /// faz a conexão com o banco de dados sql serve conforme script wen.config.
    /// </summary>
    /// <returns></returns>
    private static SqlConnection connection()
    {
        try
        {
            string sconexao = ConfigurationSettings.AppSettings["conexaoSQL"];

            SqlConnection sqlconnection = new SqlConnection(sconexao);

            //Verifica se a conexão esta fechada.
            if (sqlconnection.State == ConnectionState.Closed)
            {
                //Abri a conexão.
                sqlconnection.Open();
            }

            //Retorna o sqlconnection.
            return sqlconnection;
        }
        catch (SqlException ex)
        {

            throw ex;
        }

    }

    /// <summary>
    /// GetData abre a conexão com o SQlServer, faz o Select e retorna um Dt
    /// </summary>
    /// <param name="cmd"></param>
    /// <returns></returns>
    private DataTable GetData(SqlCommand cmd)
    {
        DataTable dt = new DataTable();
        String strConnString = ConfigurationSettings.AppSettings["conexaoSQL"];
        SqlConnection con = new SqlConnection(strConnString);
        SqlDataAdapter sda = new SqlDataAdapter();
        cmd.CommandType = CommandType.Text;
        cmd.Connection = con;
        try
        {
            con.Open();
            sda.SelectCommand = cmd;
            sda.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            con.Close();
            sda.Dispose();
            con.Dispose();
        }
    }

    /// <summary>
    /// Exporta os dados da query para o excel
    /// </summary>
    private void ExportToExcel(string sfiltro)
    {
        string strQuery = "SELECT ";

        strQuery += "a.idSolicitacao_coleta as 'Id',a.numero_solicitacao_coleta as 'Numero', ";
        strQuery += "c.Clientes_idClientes as 'Cod.Cliente', ";
    	strQuery += "e.razao_social as 'Nome', ";
    	strQuery += "c.cep_coleta as 'CEP', ";
    	strQuery += "c.logradouro_coleta as 'Endereço', ";
    	strQuery += "c.comp_endereco_coleta as 'Complemento', ";
    	strQuery += "c.bairro_coleta as 'Bairro', ";
    	strQuery += "c.municipio_coleta as 'Município', ";
    	strQuery += "c.estado_coleta as 'Estado', ";
    	strQuery += "c.contato_coleta as 'Contato', c.ddd_resp_coleta as 'DDD', c.telefone_resp_coleta as 'Telefone', ";
    	strQuery += "c.ramal_resp_coleta as 'Ramal', ";
    	strQuery += "c.depto_resp_coleta as 'Departamento' ";
	    strQuery += "from dbo.Solicitacao_coleta as a ";
        strQuery += "left outer join solicitacao_coleta_has_clientes as c ";
        strQuery += "on c.Solicitacao_coleta_idSolicitacao_coleta = a.idSolicitacao_coleta ";
        strQuery += "left outer join Solicitacao_coleta as d on a.idSolicitacao_coleta = d.idSolicitacao_coleta ";
        strQuery += "left outer join Clientes as e on e.idClientes = c.Clientes_idClientes ";
        strQuery += "left outer join Transportadoras as f on f.idTransportadoras = e.Transportadoras_idTransportadoras ";
        strQuery += "where d.Status_coleta_idStatus_coleta = 2 ";
        strQuery += sfiltro;
        strQuery += " order by a.idSolicitacao_coleta ";

        SqlCommand cmd = new SqlCommand(strQuery);
        DataTable dt = GetData(cmd);

        //Create a dummy GridView
        GridView GridView1 = new GridView();
        GridView1.AllowPaging = false;
        GridView1.DataSource = dt;
        GridView1.DataBind();

        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition","attachment;filename=DataTable.xls");
        Response.Charset = "";
        Response.ContentType = "application/vnd.ms-excel";
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);

        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            //Apply text style to each Row
            GridView1.Rows[i].Attributes.Add("class", "textmode");
        }
        GridView1.RenderControl(hw);

        //style to format numbers to string
        string style = @"<style> .textmode { mso-number-format:\@; } </style>";
        Response.Write(style);
        Response.Output.Write(sw.ToString());
        Response.Flush();
        Response.End();
    }

}