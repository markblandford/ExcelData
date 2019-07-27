using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Globalization;

namespace ExcelData
{
    public class Import : IImport
    {
        public string WorkbookPassword { get; set; }
        public int FirstDataRow { get; set; }
        public bool IgnoreBlankRows { get; set; }
        public bool UseFirstRowHeaders { get; set; }
        private bool _decrypted;
        public bool DeleteDecryptedFile { get; set; }

        private string _filename;

        private SpreadsheetDocument _spreadsheetDoc;

        public Import()
        {
        }

        public override string ToString() { return _filename; }


        public bool OpenSpreadsheet(string filename)
        {
            _decrypted = false;
            FirstDataRow = 2;

            _filename = filename;

            if (_filename != null)
            {
                if (File.Exists(_filename))
                {
                    if (Path.GetExtension(_filename) == ".xls")
                    {
                        _filename = DecryptSpreadsheet(_filename, "");
                    }
                    _spreadsheetDoc = SpreadsheetDocument.Open(_filename, false);

                    return true;
                }
                throw new FileNotFoundException(string.Format("File not found {0}", filename));
            }
            throw new ArgumentNullException("filename");
        }

        /// <summary> Open the spreadsheet to import data from. </summary>
        /// <param name="filename"> The workbook to open. </param>
        /// <param name="workbookPassword"> The password. </param>
        /// <param name="deleteDecryptedWorkbookAfterwards"> Whether to delete the decrypted workbook afterwards. Defaults to true. </param>
        /// <returns>True if the workbook was opened. </returns>
        public bool OpenSpreadsheet(string filename, string workbookPassword, bool deleteDecryptedWorkbookAfterwards = true)
        {
            _decrypted = false;
            FirstDataRow = 2;
            DeleteDecryptedFile = deleteDecryptedWorkbookAfterwards;

            _filename = filename;

            if (_filename != null)
            {
                if (File.Exists(_filename))
                {
                    if (workbookPassword != "" || Path.GetExtension(_filename) == ".xls")
                    {
                        _filename = DecryptSpreadsheet(_filename, workbookPassword);
                    }
                    _spreadsheetDoc = SpreadsheetDocument.Open(_filename, false);

                    return true;
                }
                throw new FileNotFoundException(string.Format("File not found {0}", _filename));
            }
            throw new ArgumentNullException("filename");
        }

        private Application _excelApp;
        private Microsoft.Office.Interop.Excel.Workbooks _excelWkbs;
        private Microsoft.Office.Interop.Excel.Workbook _excelWkb;

        private string DecryptSpreadsheet(string fileName, string workbookPassword)
        {
            _excelApp = new Application { Visible = false };

            _excelWkbs = _excelApp.Workbooks;

            _excelWkb = workbookPassword != "" ? _excelWkbs.Open(fileName, Password: workbookPassword) : _excelWkbs.Open(fileName);

            string appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Assembly.GetCallingAssembly().GetName().Name);

            var saveAsExtension = Path.GetExtension(fileName) == ".xls" ? ".xlsx" : Path.GetExtension(fileName);

            string newFileName = Path.Combine(appDir, Path.GetFileNameWithoutExtension(fileName) + "_nopass" + saveAsExtension);

            if (File.Exists(newFileName)) File.Delete(newFileName);

            _excelWkb.SaveAs(newFileName, Password: "", FileFormat: XlFileFormat.xlOpenXMLWorkbook);
            _excelWkb.Close();
            _excelApp.Quit();

            _decrypted = true;

            Marshal.ReleaseComObject(_excelWkb);
            Marshal.ReleaseComObject(_excelWkbs);
            Marshal.ReleaseComObject(_excelApp);

            return newFileName;
        }

        public List<T> GetExcelData<T>()
        {
            string shtName = GetSheetForObject(typeof(T));

            WorkbookPart wkb = _spreadsheetDoc.WorkbookPart;

            //Get a list of the sheets we want to import
            Sheet sheet = wkb.Workbook.Descendants<Sheet>().First(s => s.Name == shtName);

            WorksheetPart wks = (WorksheetPart)wkb.GetPartById(sheet.Id);

            Dictionary<string, string> columnsToRead = GetColumnsToRead(typeof(T));

            List<T> shtData = wks.GetSheetData<T>(wkb, shtName, FirstDataRow, columnsToRead, UseFirstRowHeaders);

            if (IgnoreBlankRows)
            {
                RemoveBlankRows<T>(ref shtData, columnsToRead);
            }

            return shtData;
        }

        private Dictionary<string, string> GetColumnsToRead(Type obj)
        {
            List<PropertyInfo> properties = obj.GetProperties().Where(p => Attribute.IsDefined(p, typeof(ExcelDataColumnAttribute))).ToList();

            Dictionary<string, string> result = new Dictionary<string, string>();

            foreach (PropertyInfo prop in properties)
            {
                result.Add(GetColumnForProperty(prop), prop.Name);
            }

            return result;
        }

        private void RemoveBlankRows<T>(ref List<T> data, Dictionary<string, string> columnsToRead)
        {
            bool isBlank = false;

            for (int i = data.Count - 1; i >= 0; i--)
            {
                T objProp = data[i];
                foreach (string prop in columnsToRead.Values)
                {
                    if (typeof(T).GetProperty(prop).GetValue(objProp) == null) isBlank = true;
                    else
                    {
                        isBlank = false;
                        break;
                    }
                }
                if (isBlank) data.Remove(objProp);
            }
        }

        private string GetSheetForObject(Type obj)
        {
            if (Attribute.IsDefined(obj, typeof(ExcelDataSheetNameAttribute)))
            {
                ExcelDataSheetNameAttribute attr = (ExcelDataSheetNameAttribute)Attribute.GetCustomAttribute(obj, typeof(ExcelDataSheetNameAttribute));

                return attr.SheetName;
            }
            return "";
        }

        private string GetColumnForProperty(PropertyInfo prop)
        {
            if (Attribute.IsDefined(prop, typeof(ExcelDataColumnAttribute)))
            {

                ExcelDataColumnAttribute attr = (ExcelDataColumnAttribute)Attribute.GetCustomAttribute(prop, typeof(ExcelDataColumnAttribute));

                return attr.Column;
            }
            return "";
        }

        #region cleanup
        private bool _disposed;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_excelWkb != null) Marshal.ReleaseComObject(_excelWkb);

                    if (_excelWkbs != null) Marshal.ReleaseComObject(_excelWkbs);

                    if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);

                    if (_spreadsheetDoc != null) _spreadsheetDoc.Dispose();

                    if (DeleteDecryptedFile && _decrypted) File.Delete(_filename);
                }
            }
            _disposed = true;
        }

        #endregion
    }

    static class OpenXmlExtensionMethods
    {
        public static List<T> GetSheetData<T>(this WorksheetPart wks, WorkbookPart wkb, string shtName, int firstDataRow, Dictionary<string, string> colsToRead, bool useHeaders = false)
        {
            string cellValue = "";
            int rowNo = -1;
            string col = "";

            try
            {

                List<T> data = new List<T>();

                using (OpenXmlReader xmlReader = OpenXmlReader.Create(wks))
                {
                    while (xmlReader.Read())
                    {
                        rowNo = -1;
                        //loop through the rows
                        if (xmlReader.ElementType == typeof(Row))
                        {
                            Dictionary<string, string> columnHeadings = new Dictionary<string, string>();
                            do
                            {
                                OpenXmlAttribute attri = xmlReader.Attributes.FirstOrDefault(r => r.LocalName == "r");

                                bool isRow;

                                isRow = xmlReader.HasAttributes && Int32.TryParse(attri.Value, out rowNo);

                                if (isRow)
                                {
                                    if (useHeaders)
                                    {
                                        if (rowNo == 1)
                                        {
                                            //Read the data
                                            xmlReader.ReadFirstChild();

                                            do
                                            {
                                                if (xmlReader.ElementType == typeof(Cell))
                                                {
                                                    Cell c = (Cell)xmlReader.LoadCurrentElement();

                                                    cellValue = ReadCell(wkb, c);

                                                    col = c.ColumnName().ToUpper();
                                                    if (colsToRead.ContainsKey(cellValue))
                                                    {
                                                        columnHeadings.Add(col, cellValue);
                                                    }
                                                }
                                            }
                                            while (xmlReader.ReadNextSibling());
                                        }
                                    }

                                    if (rowNo >= firstDataRow)
                                    {
                                        //Read the data
                                        xmlReader.ReadFirstChild();

                                        T rowData = (T)Activator.CreateInstance(typeof(T));
                                        PropertyInfo prop;

                                        do
                                        {
                                            if (xmlReader.ElementType == typeof(Cell))
                                            {
                                                Cell c = (Cell)xmlReader.LoadCurrentElement();

                                                cellValue = ReadCell(wkb, c);

                                                col = c.ColumnName().ToUpper();

                                                if (useHeaders)
                                                {
                                                    if (columnHeadings.ContainsKey(col)) col = columnHeadings[col];
                                                }

                                                if (colsToRead.ContainsKey(col))
                                                {
                                                    prop = typeof(T).GetProperty(colsToRead[col]);

                                                    Type t = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                                    object propVal;

                                                    if (t == typeof(DateTime))
                                                    {
                                                        if (cellValue == null) propVal = null;
                                                        else
                                                        {
                                                            double dt;
                                                            bool parsed = double.TryParse(cellValue, out dt);

                                                            if (parsed) propVal = DateTime.FromOADate(dt);
                                                            else propVal = DateTime.Parse(cellValue);


                                                        }
                                                    }
                                                    else if (cellValue != null && cellValue.Contains("E") && t == typeof(decimal))
                                                    {
                                                        //Scientific notation
                                                        //Convert to decimal
                                                        propVal = decimal.Parse(cellValue, NumberStyles.Float, CultureInfo.InvariantCulture);

                                                        propVal = Convert.ChangeType(propVal, t, CultureInfo.InvariantCulture);
                                                    }
                                                    else
                                                    {
                                                        propVal = (cellValue == null) ? null : Convert.ChangeType(cellValue, t, CultureInfo.InvariantCulture);
                                                    }

                                                    prop.SetValue(rowData, propVal, null);
                                                }
                                            }
                                        }
                                        while (xmlReader.ReadNextSibling());

                                        data.Add(rowData);
                                    }
                                }
                            }
                            while (xmlReader.ReadNextSibling());
                            break;
                        }
                    }

                    if (xmlReader.ElementType != typeof(DocumentFormat.OpenXml.Spreadsheet.Worksheet)) xmlReader.Skip();
                }

                return data;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Error getting sheet data [{0}, COLUMN:{1}, ROW:{2}, Value:{3}]", shtName, col, rowNo, cellValue), ex);
            }
        }

        private static string ReadCell(WorkbookPart wkb, Cell c)
        {
            string cellValue;

            if (c.CellValue == null)
            {
                cellValue = null;
            }
            else if (c.DataType != null && c.DataType == CellValues.SharedString)
            {
                SharedStringItem ssi = wkb.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));

                cellValue = ssi.Text.Text;
            }
            else cellValue = c.CellValue.InnerText;

            return cellValue;
        }

        public static string ColumnName(this Cell cell)
        {
            Regex regex = new Regex("[A-Za-z]+");

            Match match = regex.Match(cell.CellReference);

            return match.Value;
        }
    }
}
