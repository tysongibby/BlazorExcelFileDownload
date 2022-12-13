using OfficeOpenXml;
using Microsoft.JSInterop;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using System.IO;

namespace ExcelFileDownload.Pages
{
    public class ExportData
    {
        public void GenerateExcelEpPlus(IJSRuntime iJsRuntime)
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Create Excel Workbook with data
            using (var package = new ExcelPackage())
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");

                #region Header
                workSheet.Cells[1, 1].Value = "Student name";
                workSheet.Cells[1, 1].Style.Font.Size = 12;
                workSheet.Cells[1, 1].Style.Font.Bold = true;
                workSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[1, 2].Value = "Student Id";
                workSheet.Cells[1, 2].Style.Font.Size = 12;
                workSheet.Cells[1, 2].Style.Font.Bold = true;
                workSheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                #region Record
                workSheet.Cells[2, 1].Value = "Verl";
                workSheet.Cells[2, 1].Style.Font.Size = 12;
                workSheet.Cells[2, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[2, 2].Value = "1000";
                workSheet.Cells[2, 2].Style.Font.Size = 12;
                workSheet.Cells[2, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                #region Record
                workSheet.Cells[3, 1].Value = "Bertha";
                workSheet.Cells[3, 1].Style.Font.Size = 12;
                workSheet.Cells[3, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[3, 2].Value = "1001";
                workSheet.Cells[3, 2].Style.Font.Size = 12;
                workSheet.Cells[3, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                #region Record
                workSheet.Cells[4, 1].Value = "Callithria";
                workSheet.Cells[4, 1].Style.Font.Size = 12;
                workSheet.Cells[4, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[4, 2].Value = "1002";
                workSheet.Cells[4, 2].Style.Font.Size = 12;
                workSheet.Cells[4, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                // convert Excel Workbook to ByteArray
                fileContents = package.GetAsByteArray();
            }

            // Export as Excel Workbook via JavaScript
            iJsRuntime.InvokeAsync<ExportData>(
                    "saveAsFile",
                    "Student List - EPPlus.xlsx",
                    Convert.ToBase64String(fileContents)
            );
        }

        public void GenerateExcelClosedXml(IJSRuntime iJsRuntime)
        {
            // Create DataTable with data
            DataTable dt = new DataTable("Student List");
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));                        
            dt.Rows.Add(1000, "Verl");
            dt.Rows.Add(1001, "Bertha");
            dt.Rows.Add(1003, "Callithria");
            dt.Rows.Add(1004, "Gandalf");

            // Create Excel Workbook and Worksheet from DataTable
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Sheet1");

            // Convert Excel Workbook to ByteArray
            byte[] fileContents;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                fileContents= memoryStream.ToArray();

            }

            // Export as Excel Workbook via JavaScript
            iJsRuntime.InvokeAsync<ExportData>(
                "saveAsFile",
                "Student List - ClosedXML.xlsx",
                Convert.ToBase64String(fileContents)
            );


        }

    }
}
