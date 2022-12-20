using Microsoft.JSInterop;
using System;
using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;

namespace PDMSS.Services
{
    public class ExportDataFile 
    {
        /// <summary>
        /// Exports ClosedXML Excel Workbook as a browser client download
        /// </summary>
        /// <param name="iJsRuntime">Interop JavaScript Runtime or IJSRuntime</param>
        /// <param name="workbook">ClosedXML IXLWorkbook</param>
        /// <param name="fileName">String of filename to be exported without .xlsx extension</param>
        /// <returns>Async Task</returns>
        public static async Task ExportToFileAsync(IJSRuntime iJsRuntime, IXLWorkbook workbook, string fileName)
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
            await iJsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                fileName + ".xlsx",
                fileTypeString,
                Convert.ToBase64String(fileByteArray)
            );
        }
    }
}
