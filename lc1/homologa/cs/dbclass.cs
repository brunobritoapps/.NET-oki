using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// Summary description for dbclass
/// </summary>
/// 
namespace dbclass
{

    public class dbclass
    {
        public dbclass()
        {
            //
            // TODO: Add constructor logic here
            //

        }

        private static SqlConnection connection()
        {
            try
            {
                SqlConnection sqlconnection = new SqlConnection(ConfigurationSettings.AppSettings["ConnectionString"].ToString());
                //Verifica se a conexão esta fechada.
                if (sqlconnection.State == ConnectionState.Closed)
                {
                    //Abri a conexão.
                    sqlconnection.Open();
                }

                //Retorna o sqlconnection.
                return sqlconnection;
            }
            catch (Exception)
            {

                throw;
            }

        }

        /// <summary>
        /// Método que retorna um datareader com o resultado da query.
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public  SqlDataReader retornaQuery(string query)
        {
            try
            {
                //Instância o sqlcommand com a query sql que será executada e a conexão.
                SqlCommand comando = new SqlCommand(query, connection());

                //Executa a query sql.
                SqlDataReader retornaQuery = comando.ExecuteReader();

                //Fecha a conexão.
                connection().Close();

                //Retorna o dataReader com o resultado
                return retornaQuery;

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
    }
}