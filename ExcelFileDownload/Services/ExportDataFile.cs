using OfficeOpenXml;
using Microsoft.JSInterop;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;
using System.IO;

namespace ExcelFileDownload.Services
{
    public class ExportDataFile
    {
        public void ExportExcelEpPlus(IJSRuntime iJsRuntime, ExcelPackage excelPackage)
        {
            byte[] fileContents;
            string fileTypeString = "data: application / vnd.openxmlformats - officedocument.spreadsheetml.sheet; base64";

            // convert Excel Workbook to ByteArray
            fileContents = excelPackage.GetAsByteArray();

            // Export as Excel Workbook via JavaScript
            iJsRuntime.InvokeAsync<ExportDataFile>(
                    "DownloadFile",
                    "Student List - EPPlus.xlsx",
                    fileTypeString,
                    Convert.ToBase64String(fileContents)
            );
        }

        public void ExportExcelClosedXml(IJSRuntime iJsRuntime, IXLWorkbook workbook)
        {

            // Convert Excel Workbook to ByteArray
            byte[] fileByteArray;
            string fileTypeString = "data: application / vnd.openxmlformats - officedocument.spreadsheetml.sheet; base64";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                fileByteArray= memoryStream.ToArray();
            }

            // Export as Excel Workbook via JavaScript
            iJsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                "Student List - ClosedXML.xlsx",
                fileTypeString,
                Convert.ToBase64String(fileByteArray)
            );


        }

    }
}
