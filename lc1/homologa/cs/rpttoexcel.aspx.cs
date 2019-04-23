using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

public partial class rpttoexcel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String sid = Request.QueryString["id"];
        String sgrupo = Request.QueryString["grupo"];
        String sdataini = Request.QueryString["dataini"];
        String sdatafinal = Request.QueryString["datafinal"];
        String sstatus = Request.QueryString["status"];
        String sql = Request.QueryString["query"];

        //this.ExportToExcel(sid, sgrupo, sdataini, sdatafinal, sstatus, sql);
        ExportToExcel2(sid, sgrupo, sdataini, sdatafinal, sstatus, sql);
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
        Response.AddHeader("content-disposition", "attachment;filename=DataTable.xls");
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

    public void ExportToExcel2(string sid, string sgrupo, string sdataini, string sdatafinal, string sstatus, string sql)
    {
        string strQuery = sql;

        SqlCommand cmd = new SqlCommand(strQuery);
        DataTable dt = GetData(cmd);

        string str = GetTableContent(dt);

        byte[] bytes = new byte[str.Length * sizeof(char)];
        System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);

        try
        {
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment;filename=\"DataTable.csv\"");
            Response.AddHeader("Content-Length", bytes.Length.ToString());
            Response.ContentType = "application/octet-stream";
            Response.BinaryWrite(bytes);
            Response.Flush();
        }
        catch (Exception ex)
        {
            Response.ContentType = "application/vnd.ms-excel";
            Response.Write(ex.Message);
        }
        finally
        {
            Response.End();
        }
    }

    private string GetTableContent (DataTable dt)
    {
        string rw = "";
        StringBuilder builder = new StringBuilder();

        for (int a = 0; a < dt.Columns.Count; a++)
        {
            rw = GetStringNoAccents(dt.Columns[a].ToString());
            if (rw.Contains(";")) rw = "\"" + rw + "\"";
            builder.Append(rw + ";");
        }
		
		builder.Append(Environment.NewLine);

        foreach(DataRow dr in dt.Rows)
        {
            for(int i = 0; i < dt.Columns.Count; i++)
            {
                rw = dr[i].ToString();
                if (rw.Contains(";")) rw = "\"" + rw + "\"";
                builder.Append(rw + ";");
            }
            builder.Append(Environment.NewLine);
        }
        return builder.ToString();
    }

    public static string GetStringNoAccents(string str)
    {
        /** Troca os caracteres acentuados por não acentuados **/
        string[] acentos = new string[] { "ç", "Ç", "á", "é", "í", "ó", "ú", "ý", "Á", "É", "Í", "Ó", "Ú", "Ý", "à", "è", "ì", "ò", "ù", "À", "È", "Ì", "Ò", "Ù", "ã", "õ", "ñ", "ä", "ë", "ï", "ö", "ü", "ÿ", "Ä", "Ë", "Ï", "Ö", "Ü", "Ã", "Õ", "Ñ", "â", "ê", "î", "ô", "û", "Â", "Ê", "Î", "Ô", "Û" };
        string[] semAcento = new string[] { "c", "C", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "Y", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U", "a", "o", "n", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "A", "O", "N", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U" };

        for (int i = 0; i < acentos.Length; i++)
        {
            str = str.Replace(acentos[i], semAcento[i]);
        }

        return str;


        //string pattern = @"(?i)[^0-9a-záéíóúàèìòùâêîôûãõç\s]";
        //string replacement = "";
        //Regex rgx = new Regex(pattern);
        //return rgx.Replace(str, replacement);
    }
}