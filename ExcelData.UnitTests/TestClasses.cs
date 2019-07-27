using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData;

namespace ExcelData.UnitTests
{
    [ExcelDataSheetName("Security Distribution")]
    internal class Valuation
    {
        [ExcelDataColumn("Q")]
        public DateTime? ValuationDate { get; set; }
        [ExcelDataColumn("A")]
        public string AccountNumber { get; set; }
        [ExcelDataColumn("B")]
        public string AccountName { get; set; }
        [ExcelDataColumn("C")]
        public string Sedol { get; set; }
        [ExcelDataColumn("D")]
        public string Isin { get; set; }
        [ExcelDataColumn("E")]
        public string SecurityNumber { get; set; }
        [ExcelDataColumn("F")]
        public string Description { get; set; }
        [ExcelDataColumn("G")]
        public decimal? SharesPar { get; set; }
        [ExcelDataColumn("H")]
        public decimal? TradedMarketValue { get; set; }
        [ExcelDataColumn("I")]
        public decimal? Price { get; set; }
        [ExcelDataColumn("J")]
        public decimal? ExchangeRate { get; set; }
        [ExcelDataColumn("K")]
        public decimal? TradedMvAccruedIncmBase { get; set; }
        [ExcelDataColumn("L")]
        public decimal? TradedMarketValueBase { get; set; }
        [ExcelDataColumn("M")]
        public string AssetGroup { get; set; }
        [ExcelDataColumn("N")]
        public string TradingCcy { get; set; }
        [ExcelDataColumn("O")]
        public decimal? UnrealizedGainLoss { get; set; }
        [ExcelDataColumn("P")]
        public int? SecurityUniqueQual { get; set; }
        [ExcelDataColumn("R")]
        public string StockTicker { get; set; }
    }

    /// <summary> 
    /// Partial class representing a transaction.
    /// This part extends the properties from the class generated from the database. 
    /// </summary>
    [ExcelDataSheetName("Transaction Analysis")]
    internal partial class Trade
    {
        public DateTime? ValuationDate { get; set; }
        [ExcelDataColumn("Y")]
        public string AccountNumber { get; set; }
        [ExcelDataColumn("A")]
        public string Sedol { get; set; }
        [ExcelDataColumn("B")]
        public string Isin { get; set; }
        [ExcelDataColumn("C")]
        public string Description { get; set; }
        [ExcelDataColumn("D")]
        public string TranCode { get; set; }
        [ExcelDataColumn("E")]
        public string TranDescription { get; set; }
        [ExcelDataColumn("F")]
        public string TradingCurrency { get; set; }
        [ExcelDataColumn("G")]
        public string SettleCurrency { get; set; }
        [ExcelDataColumn("H")]
        public string MemoNumber { get; set; }
        [ExcelDataColumn("I")]
        public DateTime? EffectiveDate { get; set; }
        [ExcelDataColumn("J")]
        public DateTime? TradeDate { get; set; }
        [ExcelDataColumn("K")]
        public DateTime? ActualSettleDate { get; set; }
        [ExcelDataColumn("L")]
        public decimal? SharesPar { get; set; }
        [ExcelDataColumn("M")]
        public decimal? PrincipalBase { get; set; }
        [ExcelDataColumn("N")]
        public decimal? Principal { get; set; }
        [ExcelDataColumn("O")]
        public decimal? IncomeEqualization { get; set; }
        [ExcelDataColumn("P")]
        public decimal? Income { get; set; }
        [ExcelDataColumn("Q")]
        public decimal? IncomeBase { get; set; }
        [ExcelDataColumn("R")]
        public decimal? Price { get; set; }
        [ExcelDataColumn("S")]
        public decimal? PriceBase { get; set; }
        [ExcelDataColumn("T")]
        public decimal? PreviousPriceBase { get; set; }
        [ExcelDataColumn("U")]
        public string PriceCode { get; set; }
        [ExcelDataColumn("V")]
        public decimal? RealizedGainLoss { get; set; }
        [ExcelDataColumn("W")]
        public string AccountName { get; set; }
        [ExcelDataColumn("X")]
        public string Broker { get; set; }
        [ExcelDataColumn("Z")]
        public decimal? Commission { get; set; }
        [ExcelDataColumn("AA")]
        public decimal? CommissionBase { get; set; }
        [ExcelDataColumn("AB")]
        public string AssetGroup { get; set; }
        [ExcelDataColumn("AC")]
        public decimal? TotalUnrealizedFxGainLoss { get; set; }
    }

    [ExcelDataSheetName("Fund Trend Base")]
    internal class FundTrendStage
    {
        [ExcelDataColumn("D")]
        public DateTime? ValuationDate { get; set; }
        [ExcelDataColumn("A")]
        public string AccountNumber { get; set; }
        [ExcelDataColumn("C")]
        public int? AccountClass { get; set; }
        [ExcelDataColumn("E")]
        public decimal? ShareholderEquity { get; set; }
        [ExcelDataColumn("F")]
        public decimal? QuotedNav { get; set; }
        [ExcelDataColumn("G")]
        public decimal? NetIncomeAmount { get; set; }
        [ExcelDataColumn("H")]
        public decimal? CapitalBidPrice { get; set; }
        [ExcelDataColumn("I")]
        public decimal? EqualizationRate { get; set; }
        [ExcelDataColumn("J")]
        public decimal? UnrealSecurityGainLoss { get; set; }
        [ExcelDataColumn("K")]
        public decimal? UnrealizedExchangeGainLoss { get; set; }
        [ExcelDataColumn("L")]
        public decimal? UnrealFfxEquityGainLoss { get; set; }
        [ExcelDataColumn("M")]
        public decimal? UnitsInIssue { get; set; }
    }

    [ExcelDataSheetName("Delta")]
    internal class Delta
    {
        [ExcelDataColumn("B")]
        public string AssetId { get; set; }
        [ExcelDataColumn("C")]
        public string Isin { get; set; }
        [ExcelDataColumn("D")]
        public string Issue { get; set; }
        [ExcelDataColumn("E")]
        public string Ticker { get; set; }
        [ExcelDataColumn("F")]
        public string MaturityBand { get; set; }
        [ExcelDataColumn("G")]
        public string Tier { get; set; }
        [ExcelDataColumn("H")]
        public decimal? ZSpread { get; set; }
        [ExcelDataColumn("I")]
        public decimal? Dts { get; set; }
        [ExcelDataColumn("J")]
        public double? Nominal { get; set; }
        [ExcelDataColumn("K")]
        public double? Pv { get; set; }
    }

    [ExcelDataSheetName("I Don't Exist")]
    internal class InvalidSheet
    {
        [ExcelDataColumn("A")]
        public string DoesntMatter { get; set; }
    }

    [ExcelDataSheetName("Delta")]
    internal class TypeMisMatcherNull
    {
        [ExcelDataColumn("I")]
        public decimal Dts { get; set; }
    }

    [ExcelDataSheetName("Delta")]
    internal class TypeMisMatcher
    {
        [ExcelDataColumn("B")]
        public decimal ShouldBeString { get; set; }
    }

    [ExcelDataSheetName("CellTextFormat")]
    internal class TextFormat
    {
        [ExcelDataColumn("A")]
        public decimal? TextCellAsDecimal { get; set; }
        [ExcelDataColumn("B")]
        public DateTime? TextCellAsDate { get; set; }
    }
}
