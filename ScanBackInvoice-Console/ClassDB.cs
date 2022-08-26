using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScanBackInvoice_Console
{
    public class ClassDB
    {

        public DataTable SelectQueryNoLock(string query, string conn) //without transaction
        {
            DataTable dt_result = new DataTable();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(dt_result);
            }
            catch (Exception ex)
            {
                //throw;
            }
            finally
            {
                _conn.Close();
            }

            return dt_result;

        }
        public DataSet SelectQueryNoLocks(string query, string conn) //without transaction
        {
            DataSet ds_result = new DataSet();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(ds_result);
            }
            catch (Exception ex)
            {

                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return ds_result;

        }

        public bool ExecQueryNoLock(string query, string conn) //without transaction
        {
            bool result = true;
            query = " BEGIN TRY BEGIN TRAN " + query + "  COMMIT END TRY BEGIN CATCH ROLLBACK END CATCH  ";
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                cmd.ExecuteNonQuery();

            }
            catch (Exception)
            {
                result = false;

                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return result;

        }

        public DataTable ExecStoreProcNoLock(string query, string conn) //without transaction
        {
            DataTable dt_result = new DataTable();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt_result);
            }
            catch (Exception)
            {


                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return dt_result;

        }



    }
}
