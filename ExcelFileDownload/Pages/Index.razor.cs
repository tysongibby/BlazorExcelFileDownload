using ClosedXML.Excel;
using ExcelFileDownload.Services;
using Microsoft.AspNetCore.Components;
using System.Data;
using System.Threading.Tasks;
using Microsoft.JSInterop;

namespace ExcelFileDownload.Pages
{
    public partial class Index
    {
        [Inject] 
        private IJSRuntime JsRuntime { get; set; }

        [Inject]
        private NavigationManager NavigationManager { get; set; }

        // Create data table and workbook
        public DataTable DataTable { get; set; } = new DataTable("Student List");        
        private readonly XLWorkbook _workbook = new XLWorkbook();

        // Create ExportDataFile service
        private readonly ExportDataFile _export = new ExportDataFile();

        protected override void OnInitialized()
        {
            // Add data to DataTable
            DataTable.Columns.Add("Id", typeof(int));
            DataTable.Columns.Add("Name", typeof(string));
            DataTable.Rows.Add(1000, "Verl");
            DataTable.Rows.Add(1001, "Bertha");
            DataTable.Rows.Add(1002, "Callithria");
            DataTable.Rows.Add(1003, "Gandalf");

            // Create Excel Workbook and Worksheet from DataTable
            _workbook.Worksheets.Add(DataTable, "Freshman");
        }

        private async Task ExportToExcel()
        {          
            await _export.ExportToFileAsync(JsRuntime, _workbook, "StudentList", MimeType.Excel);
            NavigationManager.NavigateTo(NavigationManager.Uri, forceLoad: true);
        }

        private async Task ExportToCsv()
        {
            await _export.ExportToFileAsync(JsRuntime, _workbook, "StudentList", MimeType.Csv);
            NavigationManager.NavigateTo(NavigationManager.Uri, forceLoad: true);
        }
    }
}