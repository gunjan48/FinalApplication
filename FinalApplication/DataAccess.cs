using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;

namespace FinalApplication
{
    class DataAccess
    {
        public int executedata(string query)
        {
            try
            {
                cmd.CommandText = query;
                cmd.Connection = con;
                con.Open();
                int i = cmd.ExecuteNonQuery();
                con.Close();
                return i;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }
        OracleConnection con;
        public OracleCommand cmd;
        string sql;
        //DataTable dtc = new DataTable();

        public DataAccess()
        {
            con = new OracleConnection("Data Source=XE;  User ID=PLANTDATA; pwd=TATA");
            cmd = new OracleCommand();
            cmd.Connection = con;
        }
    }
}
