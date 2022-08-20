using System.Diagnostics;

namespace ReportExcel.Class
{
    class Close
    {
        /// <summary>
        /// <para>Метод, реализующий защиту от дурака. Закрывает все файлы Excel,
        /// чтобы можно изменить отчет</para>
        /// <para>Method protection from a fool. Close all Excel file that it can update report</para>
        /// </summary>
        internal static void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
    }
}

