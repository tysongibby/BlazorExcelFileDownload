using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Mvc;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Data;
using System.Formats.Asn1;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;

namespace ExcelFileDownload.Services
{
    public class ExportDataFile
    {
        /// <summary>
        /// Exports an IXLWorkbook as a file of designated mime type (file type) to browser client to be saved locally.
        /// </summary>
        /// <param name="jsRuntime">Interop JavaScript Runtime or IJSRuntime</param>
        /// <param name="workbook">IXLWorkbook to be exported.</param>
        /// <param name="fileName">String of filename to be exported without .xlsx extension</param>
        /// <param name="mimeType">Type of file format to be exported. Current options are .csv and .xlsx .</param>
        /// <returns>Async Task - The provided IXLWorkbook is exported to browser client to be saved.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public async Task ExportToFileAsync(IJSRuntime jsRuntime, IXLWorkbook workbook, string fileName,
            MimeType mimeType)
        {
            if (jsRuntime is null)
            {
                throw new ArgumentNullException(nameof(jsRuntime));
            }

            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            // TODO: add more mime types/file types as needed
            switch (mimeType)
            {
                case MimeType.Excel:
                    await ExportAsExcelAsync(jsRuntime, workbook, fileName, mimeType);
                    break;
                case MimeType.Csv:
                    await ExportAsCsvAsync(jsRuntime, workbook, fileName, mimeType);
                    break;
                default:
                    throw new ArgumentException("MimeType used is not valid.", nameof(mimeType));
            }
            
        }

        public async Task ExportAsCsvAsync(IJSRuntime jsRuntime, IXLWorkbook workbook, string fileName,
            MimeType mimeType, char delimiter = ',')
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                // Set file extension and mime type
                var extension = GetFileExtension(mimeType);
                var mimeTypeString = GetMimeTypeString(mimeType);

                // Create a memory stream to hold CSV data
                using var memoryStream = new MemoryStream();
                await using var streamWriter = new StreamWriter(memoryStream);

                // Save rows from worksheet to memory stream
                foreach (var row in worksheet.Rows())
                {
                    var csvRow = string.Empty;
                    foreach (var cell in row.Cells())
                    {
                       csvRow += $"{cell.Value}{delimiter}";
                    }
                    await streamWriter.WriteLineAsync(csvRow);
                }
                await streamWriter.FlushAsync();
                memoryStream.Position = 0;

                // Convert Stream to ByteArray
                var fileByteArray = memoryStream.ToArray();

                // Download the CSV file in the client's browser
                await SendFileToBrowserAsync(jsRuntime, fileByteArray, fileName, mimeTypeString, extension);

            }
        }


        public async Task ExportAsExcelAsync(IJSRuntime jsRuntime, IXLWorkbook workbook, string fileName,
            MimeType mimeType)
        {
            // Set file extension and mime type
            string extension = GetFileExtension(mimeType);
            string mimeTypeString = GetMimeTypeString(mimeType);

            // Create memory stream to hold Excel data
            using MemoryStream memoryStream = new MemoryStream();

            // Convert Excel Workbook to ByteArray
            workbook.SaveAs(memoryStream);
            byte[] fileByteArray = memoryStream.ToArray();

            // send file to client browser via JavaScript interop
            await SendFileToBrowserAsync(jsRuntime, fileByteArray, fileName, mimeTypeString, extension);
        }

        internal async Task SendFileToBrowserAsync(IJSRuntime jsRuntime, byte[] fileByteArray, string fileName,
            string mimeTypeString, string extension)
        {
            // Export as Excel Workbook via JavaScript
            await jsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                fileName + extension,
                mimeTypeString,
                Convert.ToBase64String(fileByteArray)
            );
        }

        /// <summary>
        /// Exports List of type <typeparamref name="T"/> as an Excel file to the browser client to be saved.
        /// </summary>
        /// <param name="jsRuntime">Interop JavaScript Runtime or IJSRuntime</param>
        /// <param name="list">List to be used to create an IXLWorkbook for export as Excel file to browser client.</param>
        /// <param name="fileName">String of filename without .xlsx extension</param>
        /// <param name="mimeType">Designates mime type of file for export. Also dictates file extension used for file on export.</param>
        /// <returns>Async Task - The provided List of type <typeparamref name="T"/> is exported as an Excel file to browser client to be saved.</returns>
        public async Task ExportToFileAsync<T>(IJSRuntime jsRuntime, List<T> list, string fileName, MimeType mimeType)
        {
            if (jsRuntime is null)
            {
                throw new ArgumentNullException(nameof(jsRuntime));
            }

            if (list is null || list.Count is 0)
            {
                throw new ArgumentNullException(nameof(list));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            // Create DataTable from List
            DataTable dataTable = ListToDataTable(list);
            // Create IXLWorkbook from DataTable
            IXLWorkbook workbook = DataTableToIXLWorkbook(fileName, dataTable);

            // Convert IXLWorkbook to ByteArray
            byte[] fileByteArray;

            string mimeTypeString = GetMimeTypeString(mimeType);

            string extension = GetFileExtension(mimeType);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                fileByteArray = memoryStream.ToArray();
            }

            // Export as Excel Workbook via JSRuntime
            await jsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                fileName + extension,
                mimeTypeString,
                Convert.ToBase64String(fileByteArray)
            );
        }



        /// <summary>
        /// Creates a DataTable from a List of type <typeparamref name="T"/>; using the properties of <typeparamref name="T"/> to create the DataTable Columns and the items from List of type <typeparamref name="T"/> to create the DataTables Rows.
        /// </summary>
        /// <typeparam name="T">DataType used to create the DataTable; DataType properities are used to create the DataTable Columns.</typeparam>
        /// <param name="list">List of items to create the rows of the DataTable.</param>
        /// <returns>Returns a DataTable created from the List of type <typeparamref name="T"/></returns>
        public static DataTable ListToDataTable<T>(List<T> list)
        {
            if (list is null || list.Count is 0)
            {
                throw new ArgumentNullException(nameof(list));
            }

            DataTable dataTable = new DataTable(typeof(T).Name);

            // Create data table columns from data model properties
            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo property in properties)
            {
                dataTable.Columns.Add(property.Name);
            }

            // Create data table rows from list items
            foreach (T item in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = properties[i].GetValue(item, null);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        /// <summary>
        /// Create IXLWorkbook from a DataTable
        /// </summary>
        /// <param name="workbookName">Name of IXLWorkbook to be created.</param>
        /// <param name="dataTable">DataTable to be used to create of IXLWorkbook.</param>
        /// <returns>Returns an IXLWorkbook created from a DataTable.</returns>
        public static IXLWorkbook DataTableToIXLWorkbook(string workbookName, DataTable dataTable)
        {
            if (string.IsNullOrWhiteSpace(workbookName))
            {
                throw new ArgumentNullException(nameof(workbookName));
            }

            if (dataTable is null || dataTable.Rows.Count is 0)
            {
                throw new ArgumentNullException(nameof(dataTable));
            }

            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(dataTable, workbookName);
            return workbook;
        }

        /// <summary>
        /// Gets the mimeType string for the provided MimeType.
        /// </summary>
        public string GetMimeTypeString(MimeType mimeType)
        {
            // TODO: Add more MimeType options to mimeTypeString switch statement as needed
            string mimeTypeString = mimeType switch
            {
                MimeType.Excel => "data: application / vnd.openxmlformats - officedocument.spreadsheetml.sheet; base64",
                MimeType.Csv => "data: text / csv; base64",
                _ => throw new ArgumentException("Invalid MimeType", nameof(MimeType)),
            };
            return mimeTypeString;
        }

        /// <summary>
        /// Gets the extension string for the provided MimeType.
        /// </summary>
        public string GetFileExtension(MimeType mimeType)
        {
            string extension = mimeType switch
            {
                MimeType.Excel => ".xlsx",
                MimeType.Csv => ".csv",
                _ => throw new NotSupportedException($"The provided MimeType: {mimeType} is not supported!")
            };
            return extension;
        }


    }

    /// <summary>
    /// MimeType of file to be exported. Can be Excel or CSV.
    /// </summary>
    // TODO: Add more MimeType options to MimeType enum as needed.
    public enum MimeType
    {
        Excel,
        Csv,
    }


    // TODO: Delete temp class CsvRow
    public class CsvRow
    {
        [Index(0)]
        public DateTime TimeStamp { get; set; }
        [Index(1)]
        public string SiteUrl { get; set; }
        [Index(2)]
        public string Message { get; set; }
        [Index(3)]
        public string Level { get; set; }

        public CsvRow() { }
    }

}




