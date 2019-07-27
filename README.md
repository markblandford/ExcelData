# Excel Data

ExcelData.dll is a standalone DLL providing functionality to extract data from Excel into C# classes through the use of Class and Property Attributes. It uses the OpenXML library to extract the data and so is most efficient with the newer Excel XML formats (.xlsx for example). It will however, also work with older, .xls formats.

## Developer Usage Guide
With a reference added to ExcelData.dll in a project it can be used as follows:

1. Create a Class to represent the data of a worksheet within a workbook.
2. Decorate the Class with the `[ExcelDataSheetName]` attribute where the `SheetName` parameter is the name of the Excel Worksheet:

```
  [ExcelDataSheetName("Security Distribution")]
  Public Class Valuation
  {
  }
```


3. Add properties to the Class to represent each column you wish to import from the Excel worksheet. Each property should be decorated with the `ExcelDataColumn` attribute where the `Column` parameter is the column letter:

```
  [ExcelDataSheetName("Security Distribution")]
  internal class Valuation
  {
      [ExcelDataColumn("Q")]
      public DateTime? ValuationDate { get; set; }

      [ExcelDataColumn("A")]
      public string AccountNumber { get; set; }

      [ExcelDataColumn("B")]
      public string AccountName { get; set; }
  }
```

In combination with setting `UseFirstRowHeaders = true`, it is possible to specify column headers instead of column letters.

4. To import the Excel data, create a new instance of the `Import` object. `Import` implements `IDisposable` and it is advised to use it with a `Using` statement:

```
  using ExcelData;

  public class Test
  {
      public void ImportDataExample(string fileName)
      {
          List<Valuation> valuations = new List<Valuation>();

          using (var xl = new Import())
          {
              if (xl.OpenSpreadsheet(fileName))
              {
                  exl.IgnoreBlankRows = true;
                  valuations = exl.GetExcelData<Valuation>();
              }
          }
      }
  }
```


### Properties
|Name                 |Type     |About                                                   |
|---------------------|---------|--------------------------------------------------------|
|`DeleteDecryptedFile`|`bool`   |If true and a password was provided, the decrypted file will be deleted. This is true by default. If set to false, the decrypted workbook will be saved here: `C:\\Users\\<USERNAME>\\AppData\\Roaming\\ExcelData\\<ORIGINAL_FILENAME>_nopass.xlsx`|
|`FirstDataRow`       |`int`    |The row number of the first row of data in the worksheet. By default this is 2.|
|`IgnoreBlankRows`    |`bool`   |If true, no objects will be generated in the returned list where there were blank rows. This is false by default.|
|`UseFirstRowHeaders` |`bool`   |If true, then instead of using the column letters as defined in the ExcelDataColumn attribute, the ExcelDataColumn attribute values will be used as row headers instead. This is false by default.|
|`WorkbookPassword`   |`string` |The password to open a password protected Excel file.|
