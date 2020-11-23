using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace Move
{
    public class SqlHelper : IDisposable
    {
        #region Campos

        SqlConnection oSqlConnection;

        #endregion

        #region Propriedades
        public string server { get; set; }
        public string database { get; set; }
        public string user { get; set; }
        public string pass { get; set; }
        private string ConnectionString
        {
            get
            {
                return "Server=" + server + ";" +
                        "Database=" + database + ";" +
                        "User Id=" + user + "; " +
                        "Password=" + pass + ";";
            }
        }

        #endregion

        #region Construtor

        public SqlHelper(string Database = "PRD_CRM")
        {
            server = "10.1.0.11";
            database = Database;
            user = "DEV_crm";
            pass = "DEV_crm";

            oSqlConnection = new SqlConnection(ConnectionString);
        }

        #endregion

        #region Métodos
        public DataTable ExecuteQueryDataTable(string Query, Dictionary<string, object> Parameters = null)
        {
            SqlCommand cmd = new SqlCommand(Query, oSqlConnection);
            cmd.CommandTimeout = 18000;
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();
            try
            {
                if (Parameters != null)
                    foreach (var p in Parameters)
                        cmd.Parameters.AddWithValue(p.Key, p.Value == null ? DBNull.Value : p.Value);

                cmd.CommandType = CommandType.Text;
                da.SelectCommand = cmd;
                da.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }
        }
        public DataTable ExecuteProcedureDataTable(string Procedure, Dictionary<string, object> Parameters = null)
        {
            SqlCommand cmd = new SqlCommand(Procedure, oSqlConnection);
            cmd.CommandTimeout = 18000;
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();
            try
            {
                if (Parameters != null)
                    foreach (var p in Parameters)
                        cmd.Parameters.AddWithValue(p.Key, p.Value);

                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;
                da.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }
        }
        public DataSet ExecuteProcedureDataSet(string Procedure, Dictionary<string, object> Parameters = null)
        {
            SqlCommand cmd = new SqlCommand(Procedure, oSqlConnection);
            cmd.CommandTimeout = 18000;
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            try
            {
                if (Parameters != null)
                    foreach (var p in Parameters)
                        cmd.Parameters.AddWithValue(p.Key, p.Value);

                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }
        }
        public int ExecuteProcedureInt(string Procedure, Dictionary<string, object> Parameters = null)
        {
            SqlCommand cmd = new SqlCommand(Procedure, oSqlConnection);
            cmd.CommandTimeout = 18000;
            SqlDataAdapter da = new SqlDataAdapter();
            int id;
            try
            {
                if (Parameters != null)
                    foreach (var p in Parameters)
                        cmd.Parameters.AddWithValue(p.Key, p.Value);

                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;

                cmd.Connection.Open();
                id = (int)cmd.ExecuteScalar();
                return id;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }
        }
        public void ExecuteProcedure(string Procedure, Dictionary<string, object> Parameters = null)
        {
            SqlCommand cmd = new SqlCommand(Procedure, oSqlConnection);
            cmd.CommandTimeout = 18000;
            try
            {
                if (Parameters != null)
                    foreach (var p in Parameters)
                        cmd.Parameters.AddWithValue(p.Key, p.Value);

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
                cmd.Dispose();
            }
        }
        public void BulkInsert(string nm_tabela, DataTable dataTable)
        {

            try
            {
                oSqlConnection.Open();
                using (var bulk = new SqlBulkCopy(this.oSqlConnection))
                {
                    foreach (DataColumn column in dataTable.Columns)
                        bulk.ColumnMappings.Add(column.ColumnName, column.ColumnName);


                    bulk.BulkCopyTimeout = Int32.MaxValue;
                    bulk.DestinationTableName = nm_tabela;
                    bulk.WriteToServer(dataTable);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oSqlConnection.Close();                
            }

        }
        #endregion

        #region Dispose

        // Flag: Has Dispose already been called?
        bool disposed = false;

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        ~SqlHelper()
        {
            Dispose(false);
        }

        #endregion
    }
}