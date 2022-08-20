using System.Collections.Generic;
using ReportExcel.Class;
namespace ReportExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Close.CloseProcess();
            Database.GetDataDatabase(out List<ReportData> report);
            Excel.CreateExcel(ref report, "Отчет");
        }
    }
}
