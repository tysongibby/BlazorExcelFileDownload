using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using System.Net.Http;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Components.Routing;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.Web.Virtualization;
using Microsoft.JSInterop;
using ExcelFileDownload;
using ExcelFileDownload.Shared;
using ClosedXML.Excel;
using ExcelFileDownload.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;

namespace ExcelFileDownload.Pages
{
    public partial class Index
    {
        private void ExportExcelEPPlus()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Create Excel Workbook with data then export
            using (var excelPackage = new ExcelPackage())
            {
                // Create Excel Workbook
                var excelWorksheet = excelPackage.Workbook.Worksheets.Add("Freshman");                
                // Header
                excelWorksheet.Cells[1, 1].Value = "Student name";
                excelWorksheet.Cells[1, 1].Style.Font.Size = 12;
                excelWorksheet.Cells[1, 1].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                excelWorksheet.Cells[1, 2].Value = "Student Id";
                excelWorksheet.Cells[1, 2].Style.Font.Size = 12;
                excelWorksheet.Cells[1, 2].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                // Record
                excelWorksheet.Cells[2, 1].Value = "Verl";
                excelWorksheet.Cells[2, 1].Style.Font.Size = 12;
                excelWorksheet.Cells[2, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                excelWorksheet.Cells[2, 2].Value = "1000";
                excelWorksheet.Cells[2, 2].Style.Font.Size = 12;
                excelWorksheet.Cells[2, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                // Record
                excelWorksheet.Cells[3, 1].Value = "Bertha";
                excelWorksheet.Cells[3, 1].Style.Font.Size = 12;
                excelWorksheet.Cells[3, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                excelWorksheet.Cells[3, 2].Value = "1001";
                excelWorksheet.Cells[3, 2].Style.Font.Size = 12;
                excelWorksheet.Cells[3, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                // Record
                excelWorksheet.Cells[4, 1].Value = "Callithria";
                excelWorksheet.Cells[4, 1].Style.Font.Size = 12;
                excelWorksheet.Cells[4, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                excelWorksheet.Cells[4, 2].Value = "1002";
                excelWorksheet.Cells[4, 2].Style.Font.Size = 12;
                excelWorksheet.Cells[4, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                // Record
                excelWorksheet.Cells[4, 1].Value = "Gandalf";
                excelWorksheet.Cells[4, 1].Style.Font.Size = 12;
                excelWorksheet.Cells[4, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                excelWorksheet.Cells[4, 2].Value = "1003";
                excelWorksheet.Cells[4, 2].Style.Font.Size = 12;
                excelWorksheet.Cells[4, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                // Export data
                ExportDataFile export = new ExportDataFile();
                export.ExportToFile(iJsRuntime, excelPackage);
            }
        }

        private void ExportExcelClosedXml()
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
            export.ExportToFile(iJsRuntime, workbook);
        }
    }
}