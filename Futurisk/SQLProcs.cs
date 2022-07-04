using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Futurisk
{
    public class SQLProcs
    {
        //private string conn = @"Data Source =111.118.188.128; Initial Catalog = venturaIDP; Persist Security Info=True;User ID=venturaIDP;Password=VenTur@1@3$";
        private string conn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        public DataSet SQLExecuteDataset(string query)
        {
            SqlConnection con = new SqlConnection(conn);
            SqlCommand cmd = new SqlCommand();
            try
            {
                cmd.Connection = con;
                cmd.CommandText = query;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Dispose();
                cmd.Dispose();
            }
        }

        public string SQLExecuteString(string spName, params SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection(conn);
            SqlCommand cmd = new SqlCommand();
            SqlDataReader dr;
            string retValue = "";
            try
            {
                cmd.Connection = con;
                foreach (SqlParameter param in sqlParams)
                {
                    cmd.Parameters.Add(param);
                }
                cmd.CommandText = spName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                if (con.State.Equals(ConnectionState.Open))
                {
                    con.Close();
                }
                con.Open();
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        retValue = dr[0].ToString();
                    }
                }
                dr.Close();
            }
            catch (Exception ex)
            {
                retValue = null;
            }
            finally
            {
                con.Close();
                con.Dispose();
                cmd.Dispose();
            }
            return retValue;
        }

        public int ExecuteSQLNonQuery(string query, params SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection(conn);
            SqlCommand cmd = new SqlCommand();
            try
            {
                con.Open();
                cmd.Connection = con;
                foreach (SqlParameter param in sqlParams)
                {
                    cmd.Parameters.Add(param);
                }
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = query;
                cmd.CommandTimeout = 0;
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
                cmd.Dispose();
            }
        }

        public DataSet SQLExecuteDataset(string spName, params SqlParameter[] sqlParams)
        {
            SqlConnection con = new SqlConnection(conn);
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();
            try
            {
                cmd.Connection = con;
                foreach (SqlParameter param in sqlParams)
                {
                    cmd.Parameters.Add(param);
                }
                cmd.CommandText = spName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
