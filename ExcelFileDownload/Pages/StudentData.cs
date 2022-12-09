using OfficeOpenXml;
using Microsoft.JSInterop;
using OfficeOpenXml.Style;
using System;

namespace ExcelFileDownload.Pages
{
    public class StudentData
    {
        public void GenerateExcel(IJSRuntime iJsRuntime)
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
                                
                workSheet.Cells[4, 2].Value = "1003";
                workSheet.Cells[4, 2].Style.Font.Size = 12;
                workSheet.Cells[4, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                fileContents = package.GetAsByteArray();
            }

            iJsRuntime.InvokeAsync<StudentData>(
                    "saveAsFile",
                    "Student List.xlsx",
                    Convert.ToBase64String(fileContents)
                );
        }

    }
}
