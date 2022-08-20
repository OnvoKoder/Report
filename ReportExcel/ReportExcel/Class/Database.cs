using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ReportExcel.Class
{
    internal class Database
    {
        /// <summary>
        /// <para>Получение данных из базы данных и добавление данных в список</para>
        /// <para> Getting data from database and insert data in list</para>
        /// </summary>
        /// <param name="reports"> <para>Список класса ReportData</para> <para>List class ReportData</para></param>
        internal static void GetDataDatabase(out List<ReportData> reports)
        {
            reports = new List<ReportData>();
            DbRepository db = new DbRepository();
            string query = "exec [Get Hours Report]";
            SqlConnection sql = new SqlConnection(db.GetConnect());
            try
            {
                sql.Open();
                SqlCommand cmd = new SqlCommand(query,sql);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                    while (reader.Read())
                        reports.Add(new ReportData(reader[0].ToString(), Convert.ToDouble(reader[1].ToString()), Convert.ToDouble(reader[2].ToString()), Convert.ToDouble(reader[3].ToString())));
                sql.Close();
            }
            catch (Exception){}
        }
    }
}
