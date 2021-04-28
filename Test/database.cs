using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Test2
{
    public class database
    {
        private static readonly string connStr = ConfigurationManager.AppSettings["ConnectionString"];
        static readonly string ConnectionString = ConfigurationManager.AppSettings["key"];
        private static SqlConnection conn = new SqlConnection(connStr);

        public DataTable GetData(string SQLCommand)
        {
            DataTable output = new DataTable("StockData");

             SqlDataAdapter da = new SqlDataAdapter(SQLCommand, connStr);
            da.Fill(output);

            return output;
        }

        public void WriteData(string SQLCommand)
        {
            try
            {
                SqlCommand comm = new SqlCommand(SQLCommand, conn);
                //MessageBox.Show(SQLCommand);
                comm.ExecuteNonQuery();
                comm.Dispose();
                comm = null;

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public void openconn()
        {
            conn.Open();
        }


        public void closeconn()
        {
            conn.Close();
        }





    }
}

