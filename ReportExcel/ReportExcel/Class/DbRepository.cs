namespace ReportExcel.Class
{
    class DbRepository
    {
        private static string connection;
        /// <summary>
        /// <para>Статический конструктор создается один раз. Получает строку из метода, который находится в классе ConfigurationManager</para>
        /// <para>Static construction createing once. Get string from method which it location in class ConfigurationManager</para>
        /// </summary>
        static DbRepository()
        {
            connection = ConfigurationManager.ConnectionString();
        }
        /// <summary>
        ///<para>Делает доступной строку подключения</para>
        ///<para>Do accessfull connection string</para> 
        /// </summary>
        /// <returns><para>Строка подключения</para> <para>Connection string</para></returns>
        internal string GetConnect()
        {
            return connection;
        }
    }
}
