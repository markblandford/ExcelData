using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData;

namespace ConsoleTest
{
    class Program
    {
        public static void Main(string[] args)
        {
            List<Valuation> valuations = new List<Valuation>();
            List<Trade> trades = new List<Trade>();
            List<FundTrend> fundTrends = new List<FundTrend>();

            using (
                IImport exl = new Import())
            {
                if (exl.OpenSpreadsheet(@"S:\Other\Orange\LGIM RAFI EUR Equity Master 110416.xlsx", "Lgutm6nm", true))
                {
                    exl.IgnoreBlankRows = true;
                    valuations = exl.GetExcelData<Valuation>();

                    //exl.UseFirstRowHeaders = true;
                    trades = exl.GetExcelData<Trade>();

                    fundTrends = exl.GetExcelData<FundTrend>();
                }
            }

            int valCounter = -1;
            int tradeCounter = -1;
            int fTCounter = -1;

            valCounter = valuations.ToList().Count;
            tradeCounter = trades.ToList().Count;
            fTCounter = fundTrends.ToList().Count;
        }
    }
}
