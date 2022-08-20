using System;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml;
using System.IO;

namespace ReportExcel.Class
{
    class Excel
    {
        /// <summary>
        /// <para>Создание файла Excel из данных имеющихся в списке</para>
        /// <para>Create excel file from data which it has list</para>
        /// </summary>
        /// <param name="report"><para>Список класса ReportData с данными</para> <para>List class ReportData with data</para></param>
        /// <param name="nameWorksheet"><para>Название рабочего листа</para> <para>Name worksheet</para></param>
        internal static void CreateExcel(ref List<ReportData> report,string nameWorksheet)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string dir=Directory.GetCurrentDirectory();
            string path = dir + @"/report.xlsx";
            FileInfo excelFile = new FileInfo(path);
            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                excel.Workbook.Worksheets.Delete(nameWorksheet);
                excel.Workbook.Worksheets.Add(nameWorksheet);
                var headerRow = new List<string[]>()
                    {
                        new string[] { "Дата", "Максимальное значение", "Среднее значение", "Минимальное значение "}
                    };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                var worksheet = excel.Workbook.Worksheets[nameWorksheet];
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[2, 1].LoadFromCollection(report);
                ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.ColumnClustered);
                chart.XAxis.Title.Text = "Время";
                chart.XAxis.Title.Font.Size = 10;
                chart.YAxis.Title.Text = "H";
                chart.YAxis.Title.Font.Size = 10;
                chart.SetSize(4000, 1000);
                chart.SetPosition(8, 0, 0, 0);
                var consumptionCurrentYearSeries = chart.Series.Add("B2:B52", "A2:A52");
                consumptionCurrentYearSeries.Header = "Max";
                consumptionCurrentYearSeries = chart.Series.Add("C2:C52", "A2:A52");
                consumptionCurrentYearSeries.Header = "Avg";
                consumptionCurrentYearSeries = chart.Series.Add("D2:D52", "A2:A52");
                consumptionCurrentYearSeries.Header = "Min";
                excel.SaveAs(excelFile);
            }
        }
    }
}
