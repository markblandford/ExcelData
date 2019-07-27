using System;
using System.Collections.Generic;

namespace ExcelData
{
    public interface IImport : IDisposable
    {
        /// <summary> The password, if any, to open the workbook. </summary>
        string WorkbookPassword { get; set; }
        /// <summary> The row number that contains the first row of data. Defaults to 2.</summary>
        int FirstDataRow { get; set; }
        /// <summary> Whether to ignore blank rows or not. </summary>
        bool IgnoreBlankRows { get; set; }
        /// <summary> The attributes have been set with headers instead of the column letters. </summary>
        bool UseFirstRowHeaders { get; set; }
        /// <summary> Delete the decrypted file after use. Default is true. </summary>
        bool DeleteDecryptedFile { get; set; }

        string ToString();

        /// <summary> Open the spreadsheet to import data from. </summary>
        /// <param name="filename"> The workbook to open. </param>
        /// <returns>True if the workbook was opened. </returns>
        bool OpenSpreadsheet(string filename);

        /// <summary> Open the spreadsheet to import data from. </summary>
        /// <param name="filename"> The workbook to open. </param>
        /// <param name="workbookPassword"> The password. </param>
        /// <param name="deleteDecryptedWorkbookAfterwards"> Whether to delete the decrypted workbook afterwards. Defaults to true. </param>
        /// <returns>True if the workbook was opened. </returns>
        bool OpenSpreadsheet(string filename, string workbookPassword, bool deleteDecryptedWorkbookAfterwards = true);

        /// <summary> The data returned. </summary>
        /// <typeparam name="T">The type / worksheet of the data to read.</typeparam>
        /// <returns>A list of type T.</returns>
        List<T> GetExcelData<T>();
    }
}