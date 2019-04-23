using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

public partial class rptcoletascli : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String sid = Request.QueryString["id"];
        String sgrupo = Request.QueryString["grupo"];
        String sdataini = Request.QueryString["dataini"];
        String sdatafinal = Request.QueryString["datafinal"];
        String sstatus = Request.QueryString["status"];
        String sql = Request.QueryString["query"];

        this.ExportToExcel(sid, sgrupo, sdataini, sdatafinal, sstatus, sql);

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
            Response.Write(ex);
        }
        finally
        {
            con.Close();
            sda.Dispose();
            con.Dispose();
        }
        return dt;
    }

    /// <summary>
    /// Exporta os dados da query para o excel
    /// </summary>
    private void ExportToExcel(string sid, string sgrupo, string sdataini, string sdatafinal, string sstatus,string sql)
    {
        string strQuery = sql;

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