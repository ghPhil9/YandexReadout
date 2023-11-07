using System;
using ClosedXML.Excel;

namespace YandexReadout
{
    internal class Excel
    {
        internal Excel() => Create();

        private XLWorkbook workbook;
        private IXLWorksheet worksheet;
        private int lastRow = 2;

        private void Create()
        {
            // Table
            workbook = new XLWorkbook();
            worksheet = workbook.AddWorksheet();

            // Style
            worksheet.Columns().Width = 60;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.Yellow;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Headers
            worksheet.Cell(1, 1).Value = "Поисковой запрос";
            worksheet.Cell(1, 2).Value = "Заголовок";
            worksheet.Cell(1, 3).Value = "Ссылка";
        }

        internal void NewLine(string title, string link, string request)
        {
            if (request != null) worksheet.Cell(lastRow, 1).Value = request;
            worksheet.Cell(lastRow, 2).Value = title;
            worksheet.Cell(lastRow, 3).Value = link;
            lastRow++;
        }

        internal void Save()
        {
            using (workbook)
            {
                string path = $"{Environment.CurrentDirectory}\\YandexReadout {DateTime.Now.ToString("G").Replace(":", "_")}.xlsx";
                workbook.SaveAs(path);
            }
        }
    }
}
