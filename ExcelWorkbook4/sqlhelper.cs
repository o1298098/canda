using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ExcelWorkbook4
{
    class Sqlhelper
    {
        string connectionString = "Data Source=120.76.230.35;Initial Catalog=CMA;Persist Security Info=True;User ID=sa;Password=ErgoChef@2017;";
        /// <summary>
        /// 执行sql语句
        /// </summary>
        /// <param name="sql"></param>
        public void sqlExecute(string sql)
        {
            SqlConnection con = new SqlConnection(connectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        /// <summary>
        /// 填充数据表
        /// </summary>
        public DataTable sqldataset(string sql)
        {
            SqlDataAdapter ap = new SqlDataAdapter(sql, connectionString);
            DataSet ds = new DataSet();
            ap.Fill(ds);
            ds.Tables.Add(new DataTable("Department"));
            ds.Tables[0].BeginLoadData();
            ap.Fill(ds, "Department");
            ds.Tables[0].EndLoadData();
            return ds.Tables["Department"];
        }
        /// <summary>
        /// 统计表中行数
        /// </summary>
        /// <param name="sql"></param>
        public int sqlScalar(string sql)
        {
            SqlConnection con = new SqlConnection(connectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            int i = Convert.ToInt32(cmd.ExecuteScalar());
            con.Close();
            return i;
        }
        /// <summary>
        /// 读取数据库
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public SqlDataReader sqlReader(string sql)
        {
            SqlConnection con = new SqlConnection(connectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataReader dr = cmd.ExecuteReader();
            return dr;
        }

    }
}
