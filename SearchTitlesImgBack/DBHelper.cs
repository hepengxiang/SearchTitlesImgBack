using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace SearchTitlesImgBack
{
    class DBHelper
    {
        public static string user = "";
        public static string grade = "";
        public static string constr = "Data Source=122.114.36.252,2658;Network Library = DBMSSOCN;Initial Catalog=SEARCHTITLES;User ID=utcrm;Password = 1qazse4";

        public static DataTable ExecuteQuery(string sql)
        {
            DataTable dt = new DataTable();
            try
            {

                SqlDataAdapter sda = new SqlDataAdapter(sql, constr);
                sda.Fill(dt);
            }
            catch
            {

            }
            return dt;
        }

        public static int ExecuteUpdate(string sql)
        {
            SqlConnection conn = new SqlConnection(constr);
            int i = 0;
            try
            {
                conn.Open();
                SqlCommand com = new SqlCommand(sql, conn);
                i = com.ExecuteNonQuery();
            }
            catch { }
            finally
            {
                conn.Close();
            }
            return i;
        }
    }
}
