using OfficeOpenXml;
using Microsoft.JSInterop;
using OfficeOpenXml.Style;
using System;

namespace ExcelFileDownload.Pages
{
    public class Student
    {
        public void GenerateExcel(IJSRuntime iJsRuntime)
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");

                #region Header Row
                workSheet.Cells[1, 1].Value = "Student name";
                workSheet.Cells[1, 1].Style.Font.Size = 12;
                workSheet.Cells[1, 1].Style.Font.Bold = true;
                workSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[1, 2].Value = "Student Roll";
                workSheet.Cells[1, 2].Style.Font.Size = 12;
                workSheet.Cells[1, 2].Style.Font.Bold = true;
                workSheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                #region Body 1st Row
                workSheet.Cells[1, 1].Value = "Shakib";
                workSheet.Cells[1, 1].Style.Font.Size = 12;
                workSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[1, 2].Value = "Student Roll";
                workSheet.Cells[1, 2].Style.Font.Size = 12;
                workSheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                #region Body 2nd Row
                workSheet.Cells[3, 1].Value = "Rohit";
                workSheet.Cells[3, 1].Style.Font.Size = 12;
                workSheet.Cells[3, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[3, 2].Value = "1002";
                workSheet.Cells[3, 2].Style.Font.Size = 12;
                workSheet.Cells[3, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion

                fileContents = package.GetAsByteArray();
            }

            iJsRuntime.InvokeAsync<Student>(
                    "saveAsFile",
                    "Student List.xlsx",
                    Convert.ToBase64String(fileContents)
                );
        }

    }
}
