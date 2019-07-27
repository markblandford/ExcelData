using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using NUnit.VisualStudio.TestAdapter;
using ExcelData;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelData.UnitTests
{
    [TestFixture]
    public class ImportTests
    {
        [TestCase(null, typeof(ArgumentNullException))]
        [TestCase(@"S:\Other\Orange\LGIM RAFI EUR Equity Master 110416", typeof(FileNotFoundException))]
        [TestCase(@"S:\Other\Orange\LGIM RAFI EUR Equity Master 110416.xls", typeof(FileNotFoundException))]
        [Category("Bad Files")]
        public void Import_BadFileName_ExactExceptionThrown(string badFileName, Type exceptionType)
        {
            using (var imp = new Import())
            {
                Assert.Throws(exceptionType, () => imp.OpenSpreadsheet(badFileName));
            }
        }

        [TestCase(@"S:\LGIM Asia\LGIM RAFI EUR Equity Master 110416.xlsm")]
        [Category("Bad Files")]
        public void Import_NoAccessToFile_AnyExceptionThrown(string fileName)
        {
            using (var imp = new Import())
            {
                Assert.Throws(Is.InstanceOf<Exception>(), () => imp.OpenSpreadsheet(fileName));
            }
        }

        private string GetTestDataDirectory()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TestData");
        }

        private string GetPasswordProtectedTestFile()
        {
            return Path.Combine(GetTestDataDirectory(), "PasswordProtectedFile.xlsx");
        }

        private string GetNoPasswordXlsTestFile()
        {
            return Path.Combine(GetTestDataDirectory(), "NoPasswordFile.xls");
        }

        private string _correctPassword = "Lgutm6nm";

        [Test]
        [Category("Bad Password")]
        public void Import_PasswordRequiredButNotGiven_FileFormatExceptionThrown()
        {
            string fileName = GetPasswordProtectedTestFile();

            using (var imp = new Import())
            {
                Assert.Throws(typeof(FileFormatException), () => imp.OpenSpreadsheet(fileName));
            }
        }

        [Test]
        [Category("Bad Password")]
        public void Import_PasswordRequiredButWrong_COMExceptionExceptionThrown()
        {
            const string incorrectPassword = "incorrect";

            string fileName = GetPasswordProtectedTestFile();

            using (var imp = new Import())
            {
                Assert.Throws(typeof(COMException), () => imp.OpenSpreadsheet(fileName, incorrectPassword));
            }
        }

        [TestCase(true, TestName="Return List of Valuations from password protected file and Delete decrypted file")]
        [TestCase(false, TestName = "Return List of Valuations from password protected file and Do not delete decrypted file")]
        [Category("Valuations All Good")]
        public void ImportData_PasswordRequiredAndOK_ReturnsNonEmptyListOfValuations(bool deleteDecryptedFile)
        {
            string fileName = GetPasswordProtectedTestFile();

            List<Valuation> valuations = new List<Valuation>();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName, _correctPassword, deleteDecryptedFile))
                {
                    exl.IgnoreBlankRows = true;
                    valuations = exl.GetExcelData<Valuation>();
                }
            }

            Assert.AreEqual((int)163, valuations.Count);
        }

        [TestCase(12, 1826, TestName = "Correct Data Line Number Returns List of Delta")]
        [Category("XLS File All Good")]
        public void ImportData_NoPasswordXls_ReturnsNonEmptyListOfDeltas(int dataLine, int expectedCount)
        {
            string fileName = GetNoPasswordXlsTestFile();

            List<Delta> shtData = new List<Delta>();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.FirstDataRow = dataLine;
                    exl.IgnoreBlankRows = true;
                    shtData = exl.GetExcelData<Delta>();
                }
            }

            Assert.AreEqual(expectedCount, shtData.Count);
        }

        [TestCase(2, 3, TestName = "Cells are formatted as text Returns List of TextFormat")]
        [Category("XLS File All Good")]
        public void ImportData_XlsTextFormatCells_ReturnsNonEmptyListOfDeltas(int dataLine, int expectedCount)
        {
            string fileName = GetNoPasswordXlsTestFile();

            List<TextFormat> shtData = new List<TextFormat>();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.FirstDataRow = dataLine;
                    exl.IgnoreBlankRows = true;
                    shtData = exl.GetExcelData<TextFormat>();
                }
            }

            Assert.AreEqual(expectedCount, shtData.Count);
        }

        [TestCase(1, TestName = "InCorrect Data Line Number Throws ArgumentException")]
        [Category("XLS File Exceptions Expected")]
        public void ImportData_IncorrectDataStartLine_ArgumentExceptionThrown(int dataLine)
        {
            string fileName = GetNoPasswordXlsTestFile();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.FirstDataRow = dataLine;
                    exl.IgnoreBlankRows = true;
                    Assert.Throws(Is.InstanceOf<Exception>(), () => exl.GetExcelData<Delta>());
                }
            }
        }
        
        [Test]
        [Category("XLS File Exceptions Expected")]
        public void ImportData_SheetDoesNotExist_ExceptionThrown()
        {
            string fileName = GetNoPasswordXlsTestFile();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.IgnoreBlankRows = true;
                    Assert.Throws(Is.InstanceOf<Exception>(), () => exl.GetExcelData<InvalidSheet>());
                }
            }
        }

        [TestCase(TestName = "Null cell in non-null decimal property type")]
        [Category("XLS Delta File Exceptions Expected")]
        public void ImportData_NullCellTypeMismatchError_ArgumentExceptionThrown()
        {
            string fileName = GetNoPasswordXlsTestFile();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.IgnoreBlankRows = true;
                    Assert.Throws(Is.InstanceOf<Exception>(), () => exl.GetExcelData<TypeMisMatcherNull>());
                }
            }
        }
        
        [TestCase(TestName = "String cell in decimal property type")]
        [Category("XLS Delta File Exceptions Expected")]
        public void ImportData_StringCellDecimalTypeMismatchError_ArgumentExceptionThrown()
        {
            string fileName = GetNoPasswordXlsTestFile();

            using (var exl = new Import())
            {
                if (exl.OpenSpreadsheet(fileName))
                {
                    exl.IgnoreBlankRows = true;
                    Assert.Throws(Is.InstanceOf<Exception>(), () => exl.GetExcelData<TypeMisMatcher>());
                }
            }
        }
    }
}
