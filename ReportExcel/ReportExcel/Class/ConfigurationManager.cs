using System.IO;
using System.Linq;

namespace ReportExcel.Class
{
    class ConfigurationManager
    {
        /// <summary>
        /// <para>Метод читает конфигурационный файл. Слздает строку подключения для базы данных</para>
        /// <para>Method read config file. Create connection string for a database</para>
        /// </summary>
        /// <returns><para>Строку подключения</para> <para>Connection string</para></returns>
       internal static string ConnectionString()
       {
            string server="", database="", psw="", login="", signIN="";
            string dir = Directory.GetCurrentDirectory();
            string path = dir + @"/config.txt";
            server = File.ReadLines(path).Skip(0).First().Trim();
            database = File.ReadLines(path).Skip(0).ElementAt(1).Trim();
            if (File.ReadLines(path).Skip(0).ElementAt(2).Trim().StartsWith("#"))
            {
                login = File.ReadLines(path).Skip(0).ElementAt(3).Trim();
                psw = File.ReadLines(path).Skip(0).Last().Trim();
            }
            else
                signIN = File.ReadLines(path).Skip(0).ElementAt(2).Trim();
            if (signIN != null)
                return server + database + signIN;
            else return server+ database+login+psw;
       }
    }
}
