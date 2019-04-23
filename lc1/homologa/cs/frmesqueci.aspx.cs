using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

public partial class _Default : System.Web.UI.Page
{
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
    /// Método que retorna um datareader com o resultado da query.
    /// </summary>
    /// <param name="query"></param>
    /// <returns></returns>
    public string retornaQuery(string query)
    {
        try
        {
            //Instância o sqlcommand com a query sql que será executada e a conexão.
            SqlCommand comando = new SqlCommand(query, connection());

            //Executa a query sql.
            SqlDataReader retornaQuery = comando.ExecuteReader();

            if (retornaQuery.Read())
            {
                connection().Close();
                string sretorno = retornaQuery[0].ToString();
                return sretorno;
            }
            else
            {
                connection().Close();
                return "conteudo_embranco";
            }

        }
        catch (SqlException ex)
        {
            throw ex;
        }

    }

    public string retornaQuery2(string query)
    {
        try
        {
            //Instância o sqlcommand com a query sql que será executada e a conexão.
            SqlCommand comando = new SqlCommand(query, connection());

            //Executa a query sql.
            SqlDataReader retornaQuery2 = comando.ExecuteReader();

            if (retornaQuery2.Read())
            {
                connection().Close();
                string sretorno2 = retornaQuery2[1].ToString();
                return sretorno2;
            }
            else
            {
                connection().Close();
                return "conteudo_embranco";
            }

        }
        catch (SqlException ex)
        {
            throw ex;
        }

    }

    public DataSet retornaQueryDataSet(string query)
    {
        try
        {
            //Instância o sqlcommand com a query sql que será executada e a conexão.
            SqlCommand comando = new SqlCommand(query, connection());

            //Instância o sqldataAdapter.
            SqlDataAdapter adapter = new SqlDataAdapter(comando);

            //Instância o dataSet de retorno.
            DataSet dataSet = new DataSet();

            //Atualiza o dataSet
            adapter.Fill(dataSet);

            //Retorna o dataSet com o resultado da query sql.
            return dataSet;
        }
        catch (Exception ex)
        {
            throw ex;
        }

    } 

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Page.IsPostBack)
        {

            return; 
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string suser;
        string ssenha;
        string squery = "select usuario,senha from [marketingoki2].[dbo].[Contatos] where email = '" + txtBemail.Text.ToLower() + "' and status_contato = 1 ";

        suser  = retornaQuery(squery);
        ssenha = retornaQuery2(squery);

        if (IsPostBack && !string.IsNullOrEmpty(txtBemail.Text))
        {
            //crio objeto responsável pela mensagem de email
            MailMessage objEmail = new MailMessage();

            objEmail.From = new MailAddress("sustentabilidadeoki@sustentabilidadeoki.com.br");
            objEmail.To.Add(txtBemail.Text.ToLower());
            //objEmail.CC.Add("peterson.aquino@hotmail.com");
            objEmail.IsBodyHtml = true;
            objEmail.Subject = "OKIDATA - Recuperamos sua senha";

            String cBody = "";

            cBody += "<!DOCTYPE html>";
            cBody += "<html>";
            cBody += "<head>";
            cBody += "<link rel='stylesheet' type='text/css' href='css/geral.css'>";
            cBody += "<title>Recuperar Senha</title>";
            cBody += "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>";
            cBody += "</head>";
            cBody += "<table cellspacing='0' cellpadding='0' width='775'>";
            cBody += "		<tr>";
            cBody += "			<td width='11' background='img/Bg_LatEsq.gif'>&nbsp;</td>";
            cBody += "";
            cBody += "			<td id='conteudo'>";
            cBody += "				<table width='100%' cellspacing='0' cellpadding='0'>";
            cBody += "					<tr>";
            cBody += "						<td width='100%'>";
            cBody += "                          &nbsp;<img src='http://www.sustentabilidadeoki.com.br/banner_oki.png' />";
            cBody += "							&nbsp;</p>";
            cBody += "                          <font face='Verdana, Arial' color='black'  size='4'>";
            cBody += "							&nbsp;Solicita&ccedil;&atilde;o de Recupera&ccedil;&atilde;o de senha</p></Font>";
            cBody += "							<tr><td>";
            cBody += "							<font face='Verdana, Arial' color='black'  size='2'>&nbsp;Estamos reenviando sua senha do portal Sustentabilidade OKI:<br></td></tr>";
            cBody += "						</td>";
            cBody += "					</tr>";
            cBody += "					<tr>";
            cBody += "						<td width='100%'>&nbsp;<font face='Verdana, Arial' color='black'  size='1'><b>Usu&aacute;rio: " + suser + "<b></font><br>";
            cBody += "						</td>";
            cBody += "					</tr>";
            cBody += "					<tr>";
            cBody += "						<td width='100%'>&nbsp;<font face='Verdana, Arial' color='black'  size='1'><b>Senha: " + ssenha + "<b></font><br>";
            cBody += "						</td>";
            cBody += "					</tr>";
            cBody += "					<tr>";
            cBody += "						<t;d width='200'>&nbsp;<br>";
            cBody += "						</td>";
            cBody += "					</tr>					";
            cBody += "				</table>";
            cBody += "			</td>";
            cBody += "			<td width='11' background='img/Bg_LatDir.gif'>&nbsp;</td>";
            cBody += "		</tr>";
            cBody += "		</table>";
            cBody += "</html>";

            objEmail.Body = cBody;
            objEmail.BodyEncoding = Encoding.GetEncoding("ISO-8859-1");

            SmtpClient objSmtp = new SmtpClient();
            objSmtp.Host = "smtp.sustentabilidadeoki.com.br"; //"mail.okidata.com.br";
            objSmtp.Credentials = new NetworkCredential("sustentabilidadeoki@sustentabilidadeoki.com.br", "Oki!321!"); //("nfe@okidata.com.br", "!nfe321!");
            //objSmtp.Credentials = new NetworkCredential("sustentabilidadeoki@sustentabilidadeoki.com.br", "Oki!321!");
            objSmtp.Port = 587;
            objSmtp.Send(objEmail);

            Response.Redirect("frmlogincliente.asp");

        }
        else
        {
            if (string.IsNullOrEmpty(txtBemail.Text))
            {
                
            }
        }
    }
    protected void TextBox1_TextChanged(object sender, EventArgs e)
    {

    }
}