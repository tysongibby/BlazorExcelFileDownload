using ClosedXML.Excel;
using ExcelFileDownload.Services;
using System.Data;
using System.Threading.Tasks;

namespace ExcelFileDownload.Pages
{
    public partial class Index
    {
        private async Task ExportExcelClosedXml()
        {
            // Create DataTable with data
            DataTable dataTable = new DataTable("Student List");
            dataTable.Columns.Add("Id", typeof(int));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Rows.Add(1000, "Verl");
            dataTable.Rows.Add(1001, "Bertha");
            dataTable.Rows.Add(1002, "Callithria");
            dataTable.Rows.Add(1003, "Gandalf");
            // Create Excel Workbook and Worksheet from DataTable
            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(dataTable, "Freshman");
            // Export data
            ExportDataFile export = new ExportDataFile();
            await export.ExportToFileAsync(iJsRuntime, workbook, "StudentList");
        }
    }
}