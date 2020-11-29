using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace BhavCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            WebClient myWebClient = new WebClient();
            var i = 0;
            Dictionary<string, decimal?> stocksAndTheirOpenPrice = new Dictionary<string, decimal?>();
            Dictionary<string, List<NSE.StockData>> last50DaysStockData = new Dictionary<string, List<NSE.StockData>>();
            Dictionary<string, int> hasBeenPositiveFor = new Dictionary<string, int>();
            Dictionary<string, int> hasBeenNegativeFor = new Dictionary<string, int>();
            Dictionary<string, int> hasBeenZeroDiffFor = new Dictionary<string, int>();
            
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/Companies_proposed_to_be_delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx");
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.DelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();

            foreach (int item in Enumerable.Range(0, 50).ToList())
            {
                var consideredDate = DateTime.Now.Date.AddDays(-item);
                if (consideredDate.DayOfWeek != DayOfWeek.Saturday && consideredDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    var dateString = consideredDate.ToString("dd-MM-yy").Split("-");
                    var stringKey = dateString[0] + dateString[1] + dateString[2]; // Date agianst which we are interetsed to get the records.
                    var url = "https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip"; // NSE Data for a given day.
                    
                    try
                    {
                        var downloadFilePath = @"C:\Trading\BhavCopy\Last50DaysNSE\" + stringKey + ".zip"; 
                        var extractPath = @"C:\Trading\BhavCopy\NSEResponse";
                        myWebClient.DownloadFile(url, downloadFilePath);
                        try { Directory.Delete(extractPath, true); } catch { } // Clearing out NSEResponse folder after every day's stockData is added up
                        System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                        CreateXlsxFile(extractPath, stringKey); // Moving CSV File contents to XLSX format as EPPlus can only read xlsx formatted data.
                        List<NSE.StockData> stockData = new EpPlusHelper().ReadFromExcel<List<NSE.StockData>>(extractPath + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                        var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList(); // Here, we get stocks which arent delisted or on the delistable notice list.
                        last50DaysStockData.Add(stringKey, delistFilteredEquityStockData);
                    }
                    catch (Exception ex) { continue; }
                }
            }
            Dictionary<string, decimal?>[] diffOfOpenPrevCloseForEachStock = new Dictionary<string, decimal?>[last50DaysStockData.Values.Count];

            foreach (var item in last50DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null && x.OPEN_PRICE != null && x.PREV_CL_PR != null))
                {
                    var validOpenPrice = decimal.TryParse(stock.OPEN_PRICE, out decimal openPrice);
                    var validPrevClosePrice = decimal.TryParse(stock.PREV_CL_PR, out decimal prevClosePrice);
                    if (validOpenPrice && validPrevClosePrice) // If Open & PreClose prices are not null in excel sheet
                    {
                        var priceDifference = openPrice - prevClosePrice;
                        if (!stocksAndTheirOpenPrice.ContainsKey(stock.SYMBOL))
                            stocksAndTheirOpenPrice.Add(stock.SYMBOL, openPrice);
                        
                        //Below section is to check how consistently the price has been positive wrt (Open-Prev Close)
                        if (priceDifference > 0 && !hasBeenPositiveFor.ContainsKey(stock.SYMBOL))
                            hasBeenPositiveFor.Add(stock.SYMBOL, 1); // Has (Open-PrevClose) as +ve for these many days.
                        else if (priceDifference > 0)
                            hasBeenPositiveFor[stock.SYMBOL] += 1;
                        else if (priceDifference < 0 && !hasBeenNegativeFor.ContainsKey(stock.SYMBOL))
                            hasBeenNegativeFor.Add(stock.SYMBOL, 1);
                        else if (priceDifference < 0)
                            hasBeenNegativeFor[stock.SYMBOL] += 1;
                        else if(!hasBeenZeroDiffFor.ContainsKey(stock.SYMBOL))
                            hasBeenZeroDiffFor.Add(stock.SYMBOL, 1);
                        else
                            hasBeenZeroDiffFor[stock.SYMBOL]++; // Open - PrevClose = 0 for these many days.

                        // Storing the Difference of Open - Prev. Close for each day (denoted by i) and for each stock symbol into : diffOfOpenPrevCloseForEachStock
                        if (diffOfOpenPrevCloseForEachStock[i] == null)
                            diffOfOpenPrevCloseForEachStock[i] = new Dictionary<string, decimal?> { { stock.SYMBOL, priceDifference } };
                        else
                            diffOfOpenPrevCloseForEachStock[i].Add(stock.SYMBOL, priceDifference);
                    }
                }
                i++;
            }

            i = 0;
            Dictionary<string, decimal?> eachStockAverageOver50Days = new Dictionary<string, decimal?>();
            List<string> stockNames = new List<string>();
            Dictionary<string, int> counterOfIthStock = new Dictionary<string, int>();
            
            foreach (var item in last50DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null))
                {
                    var isAveragePresent = diffOfOpenPrevCloseForEachStock[i].TryGetValue(stock.SYMBOL, out decimal? openCloseDiff); // Against each day sugested by i, check if this particular stock's value exists.
                    if (isAveragePresent && openCloseDiff != null)
                    {
                        if (!stockNames.Contains(stock.SYMBOL))
                            stockNames.Add(stock.SYMBOL);

                        if (!counterOfIthStock.ContainsKey(stock.SYMBOL))
                            counterOfIthStock.Add(stock.SYMBOL, 1); // get total count to use later for diving total calculated in line 113
                        else
                            counterOfIthStock[stock.SYMBOL] += 1;

                        if (!eachStockAverageOver50Days.ContainsKey(stock.SYMBOL))
                            eachStockAverageOver50Days.Add(stock.SYMBOL, openCloseDiff);
                        else
                            eachStockAverageOver50Days[stock.SYMBOL] += openCloseDiff; // eachStockAverageOver50Days , currently add all values and save total
                    }
                }
                i++;
            }

            foreach (var name in stockNames.Where(x => eachStockAverageOver50Days.ContainsKey(x) && counterOfIthStock.ContainsKey(x)))
                eachStockAverageOver50Days[name] = eachStockAverageOver50Days[name] / counterOfIthStock[name]; // find average used in line 113 and line 106

            var numbers = new List<string>();
            foreach (var num in Enumerable.Range(0, 9).ToArray())
                numbers.Add("'" + num + "'");
            var screenedAvgOpenCloseStocks = eachStockAverageOver50Days.Where(x => stocksAndTheirOpenPrice[x.Key] != null && x.Value >= (stocksAndTheirOpenPrice[x.Key] * 0.02m)); // (> 2% of the stock's value happens to be the open prevClose difference on an average) // && stocksAndTheirOpenPrice[x.Key] > 100 && stocksAndTheirOpenPrice[x.Key] < 10000 
            var screenedStocks = from avgOpenClose in screenedAvgOpenCloseStocks
                                 join OpenPrice in stocksAndTheirOpenPrice on avgOpenClose.Key equals OpenPrice.Key
                                 where OpenPrice.Value != avgOpenClose.Value && hasBeenPositiveFor.ContainsKey(OpenPrice.Key) && (((decimal?)hasBeenPositiveFor[OpenPrice.Key] / ((decimal)last50DaysStockData.Count)) > 0.8m)// Has yielded positive outcomes for more than 80% of the time on an average               && !numbers.Contains(avgOpenPriceClose.Key.Substring(0, 1)) // && OpenPrice.Value > 160 && OpenPrice.Value < 25000 // && avgOpenPriceClose.Value < (.07m * OpenPrice.Value) // To remove fake ones which show prices for one day only and inactive on other days
                                 select new {PricePositiveForDays = hasBeenPositiveFor[OpenPrice.Key], PricePositiveByTotalDays = (((decimal?)hasBeenPositiveFor[OpenPrice.Key])/ ((decimal)last50DaysStockData.Count)), Stock = avgOpenClose.Key, AverageOpenCloseDiff = avgOpenClose.Value, OpenPrice50DaysAgo = OpenPrice.Value, OpenCloseDiffToOpenRatio = (avgOpenClose.Value / OpenPrice.Value) * 100 };
            string json = JsonConvert.SerializeObject(screenedStocks, Formatting.Indented);
            Console.WriteLine(json);
            Console.ReadLine();
        }

        private static void CreateXlsxFile(string extractPath, string stringKey)
        {
            string csvFileName = extractPath + "\\" + "Pd" + stringKey + ".csv";
            string excelFileName = extractPath + "\\" + "Pd" + stringKey + ".xlsx";
            string worksheetsName = "Pd" + stringKey;
            bool firstRowIsHeader = true;
            var format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.EOL = "\n";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                package.Save();
            }
        }
    }
}
