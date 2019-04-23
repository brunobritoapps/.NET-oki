using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

public partial class rptcoletas : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String sid = Request.QueryString["id"];
        String sgrupo = Request.QueryString["grupo"];
        String sdataini = Request.QueryString["dataini"];
        String sdatafinal = Request.QueryString["datafinal"];
        String sstatus = Request.QueryString["status"];

        this.ExportToExcel(sid, sgrupo, sdataini, sdatafinal, sstatus);

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
    private void ExportToExcel(string sid, string sgrupo, string sdataini, string sdatafinal, string sstatus)
    {
        string strQuery = "SELECT ";

        strQuery += " convert( nvarchar(10), A.[data_solicitacao], 103), A.[numero_solicitacao_coleta], C.[idClientes], C.[razao_social], C.[nome_fantasia], A.[qtd_cartuchos], convert( nvarchar(10), A.[data_programada],103) , ";
        strQuery += " D.[status_coleta] ";
        strQuery += "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A ";
        strQuery += "LEFT JOIN [marketingoki2].[dbo].[Status_coleta] as D on D.[idStatus_coleta] = A.[Status_coleta_idStatus_coleta] ";
        strQuery += "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B ";
        strQuery += "ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] ";
        strQuery += "LEFT JOIN [marketingoki2].[dbo].[Clientes] AS C ";
        strQuery += "ON B.[Clientes_idClientes] = C.[idClientes] where A.[isMaster] = 0 ";
        strQuery += "   AND C.[idClientes] = " + sid+ "";

        if ( !String.IsNullOrEmpty(sstatus.ToString()) && Convert.ToInt16(sstatus.ToString()) != 0 )
	    {
            strQuery += "   AND A.[Status_coleta_idStatus_coleta] = " + sstatus + "";
    	}
        if (!String.IsNullOrEmpty(sdataini.ToString()) && String.IsNullOrEmpty(sdatafinal.ToString()))
        {
            strQuery += "   AND A.[data_solicitacao] between convert(datetime, '" + sdataini + " 00:01') and convert(datetime, '" + sdataini + " 23:59')";            
        }
        else if (!String.IsNullOrEmpty(sdataini.ToString()) && !String.IsNullOrEmpty(sdatafinal.ToString()))
        {
            strQuery += "   AND A.[data_solicitacao] between convert(datetime, '" + sdataini + " 00:01') and convert(datetime, '" + sdatafinal + " 23:59')";
        }
        else if (String.IsNullOrEmpty(sdataini.ToString()) && !String.IsNullOrEmpty(sdatafinal.ToString()))
        {
            strQuery += "   AND A.[data_solicitacao] between convert(datetime, '" + sdatafinal + " 00:01') and convert(datetime, '" + sdatafinal + " 23:59')";
        }
        
        /*
        strQuery += " A.[idSolicitacao_coleta], A.[Status_coleta_idStatus_coleta], A.[numero_solicitacao_coleta], A.[qtd_cartuchos], A.[qtd_cartuchos_recebidos], A.[data_solicitacao] ";
        strQuery += ",A.[data_aprovacao], A.[data_envio_transportadora], A.[data_entrega_pontocoleta], A.[data_recebimento], A.[motivo_status], A.[isMaster], B.[Solicitacao_coleta_idSolicitacao_coleta] ";
        strQuery += ",B.[typeColect], B.[Pontos_coleta_idPontos_coleta], B.[Contatos_idContatos], B.[Clientes_idClientes], B.[cep_coleta], B.[logradouro_coleta] ";
        strQuery += ",B.[bairro_coleta], B.[numero_endereco_coleta], B.[comp_endereco_coleta], B.[municipio_coleta], B.[estado_coleta], B.[ddd_resp_coleta] ";
        strQuery += ",B.[telefone_resp_coleta], B.[contato_coleta], C.[idClientes], C.[Grupos_idGrupos], C.[Categorias_idCategorias], C.[razao_social] ";
        strQuery += ",C.[nome_fantasia], C.[cnpj], C.[inscricao_estadual], C.[ddd], C.[telefone], C.[compl_endereco], C.[compl_endereco_coleta], C.[numero_endereco] ";
        strQuery += ",C.[numero_endereco_coleta], C.[contato_respcoleta], C.[ddd_respcoleta], C.[telefone_respcoleta], C.[numero_sequencial], C.[data_atualizacao_sequencial] ";
        strQuery += ",C.[minCartuchos], C.[typeColect], C.[status_cliente], C.[motivo_status], C.[bonus_type], C.[Transportadoras_idTransportadoras] ";
        strQuery += ",C.[tipopessoa], C.[cod_cli_consolidador], C.[cod_bonus_cli], A.[data_programada] ";
        strQuery += "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A ";
        strQuery += "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B ";
        strQuery += "ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] ";
        strQuery += "LEFT JOIN [marketingoki2].[dbo].[Clientes] AS C ";
        strQuery += "ON B.[Clientes_idClientes] = C.[idClientes] where A.[isMaster] = 0 ";
        */
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