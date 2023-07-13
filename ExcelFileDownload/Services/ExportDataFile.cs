using ClosedXML.Excel;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelFileDownload.Services
{
    public class ExportDataFile
    {
        /// <summary>
        /// Exports an IXLWorkbook as an Excel file to browser client to be saved.
        /// </summary>
        /// <param name="iJsRuntime">Interop JavaScript Runtime or IJSRuntime</param>
        /// <param name="workbook">IXLWorkbook to be exported.</param>
        /// <param name="fileName">String of filename to be exported without .xlsx extension</param>
        /// <returns>Async Task - The provided IXLWorkbook is exported to browser client to be saved.</returns>
        public async Task ExportToFileAsync(IJSRuntime iJsRuntime, IXLWorkbook workbook, string fileName)
        {
            if (iJsRuntime is null)
            {
                throw new ArgumentNullException(nameof(iJsRuntime));
            }
            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            // Convert Excel Workbook to ByteArray
            byte[] fileByteArray;
            string fileTypeString = "data: application / vnd.openxmlformats - officedocument.spreadsheetml.sheet; base64";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                fileByteArray = memoryStream.ToArray();
            }
            // Export as Excel Workbook via JavaScript
            await iJsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                fileName + ".xlsx",
                fileTypeString,
                Convert.ToBase64String(fileByteArray)
            );
        }

        /// <summary>
        /// Exports List of type <typeparamref name="T"/> as an Excel file to the browser client to be saved.
        /// </summary>
        /// <param name="iJsRuntime">Interop JavaScript Runtime or IJSRuntime</param>
        /// <param name="list">List to be used to create an IXLWorkbook for export as Excel file to browser client.</param>
        /// <param name="fileName">String of filename without .xlsx extension</param>
        /// <returns>Async Task - The provided List of type <typeparamref name="T"/> is exported as an Excel file to browser client to be saved.</returns>
        public async Task ExportToFileAsync<T>(IJSRuntime iJsRuntime, List<T> list, string fileName)
        {
            if (iJsRuntime is null)
            {
                throw new ArgumentNullException(nameof(iJsRuntime));
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
            string fileTypeString = "data: application / vnd.openxmlformats - officedocument.spreadsheetml.sheet; base64";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                fileByteArray = memoryStream.ToArray();
            }
            // Export as Excel Workbook via JSRuntime
            await iJsRuntime.InvokeAsync<ExportDataFile>(
                "DownloadFile",
                fileName + ".xlsx",
                fileTypeString,
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


    }
}

