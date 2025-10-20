using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using HtmlToOpenXml;
//using IronPdf;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using SelectPdf;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
namespace Testing
{
    public class FactSheetData
    {
        // 1. Report Date
        public DateTime ReportDate { get; set; }

        // 3. Current AuM
        public decimal CurrentAuM_USD { get; set; }

        // 4. Fund Performance Table Data
        public List<MonthlyReturn> MonthlyPerformanceGrid { get; set; } = new List<MonthlyReturn>();

        // 5. NAV Evolution Data (Time-series for charting)
        public List<NavEvolutionEntry> NavEvolution { get; set; } = new List<NavEvolutionEntry>();

        // 6. Risk Data
        public Dictionary<string, decimal> Risks { get; set; } = new Dictionary<string, decimal>();

        // 7. Counterparty Risk Data
        public List<CounterpartyRiskEntry> CounterpartyRisks { get; set; } = new List<CounterpartyRiskEntry>();

        // 8. Assets Breakdown (for charting/table)
        public List<AssetBreakdownEntry> AssetsBreakdown { get; set; } = new List<AssetBreakdownEntry>();
    }
    public class PerformanceEntry
    {
        public string Period { get; set; } // e.g., "Month", "YTD", "Inception"
        public string ShareClass { get; set; } // e.g., "USD Share Class"
        public decimal FundReturn { get; set; } // The actual return value
        public decimal BenchmarkReturn { get; set; } // If you have a benchmark
    }
    public class NavEvolutionEntry
    {
        public DateTime Date { get; set; }
        public double NavPerShare { get; set; }
    }
    public class CounterpartyRiskEntry
    {
        public string Counterparty { get; set; } // e.g., "Prime Broker A", "Exchange B"
        public decimal ExposureUSD { get; set; }
        public decimal ExposurePercentNAV { get; set; }
    }
    public class RawCounterpartyExposure
    {
        public string ExchangeName { get; set; } // The name as it appears in the Excel file
        public decimal NAV { get; set; }
    }
    public class AssetBreakdownEntry
    {
        public string Category { get; set; } // e.g., "BTC", "ETH", "Stablecoins", "Cash"
        public decimal ValueUSD { get; set; }
        public decimal ValuePercent { get; set; }
    }
    public class MonthlyReturn
    {
        public DateTime Date { get; set; }
        public decimal Return { get; set; }
        public string MonthName { get; set; }
    }

    public class ExcelDataProcessor
    {
        private readonly string _basePath = @"H:\Shared drives\Melanion Management\Trading\Operations\Excel\NAV SMN\";
        private const string HistoricalPerformanceFilePath = @"H:\Shared drives\Melanion Management\Users\Anthony\FactSheet.xlsx";

        private string GetFilePath(DateTime reportDate)
        {
            string folderDate = reportDate.ToString("yyyy-MM");
            string fileDate = reportDate.ToString("yyyyMMdd");
            string fileName = $"ExchangeSnapshot_SigmaNeutral_{fileDate}_1300.xlsx";
            return Path.Combine(_basePath, folderDate, fileName);
        }

        public FactSheetData ExtractData(DateTime reportDate, decimal currentMonthlyReturn)
        {
            var factData = new FactSheetData
            {
                ReportDate = reportDate
            };

            string filePath = GetFilePath(reportDate);

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found at: {filePath}");
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new InvalidOperationException("No worksheet found in the Excel file.");
                }

                factData.CurrentAuM_USD = ExtractDecimalFromCell(worksheet, "L1");
                factData.NavEvolution = ExtractNavEvolution(worksheet, "D50:E100");
                factData.Risks = GetRiskScenarios(reportDate).Result;

                ExtractCounterpartyRisks(factData, worksheet);
                ExtractAssetBreakdown(factData, worksheet);
            }
            UpdateAndSaveHistoricalPerformance(reportDate, currentMonthlyReturn);
            factData.MonthlyPerformanceGrid = ExtractMonthlyPerformanceGrid(reportDate, currentMonthlyReturn);
            factData.NavEvolution = LoadNavEvolution(reportDate, (double)currentMonthlyReturn);


            return factData;
        }

        public List<NavEvolutionEntry> LoadNavEvolution(DateTime reportDate, double nextMonthGrowthRate)
        {
            var navEvolution = new List<NavEvolutionEntry>();

            using (var workbook = new XLWorkbook(HistoricalPerformanceFilePath))
            {
                var sheet = workbook.Worksheet("NAV");

                // Read existing data
                int row = 1;
                bool reportDateExists = false;
                bool isFirst = true;
                while (!sheet.Cell(row, 1).IsEmpty())
                {
                    var date = sheet.Cell(row, 1).GetDateTime();
                    var value = sheet.Cell(row, 2).GetDouble();

                    if (date.Date == reportDate.Date)
                        reportDateExists = true;

                    if (!isFirst)
                    {
                        // For all except the first value, take 1st day of the month
                        date = date.AddMonths(1);
                    }

                    navEvolution.Add(new NavEvolutionEntry
                    {
                        Date = date,
                        NavPerShare = value
                    });

                    isFirst = false;
                    row++;
                }

                if (!reportDateExists)
                {
                    // Calculate NAV based on last entry
                    var lastEntry = navEvolution.Last();
                    var nextValue = lastEntry.NavPerShare * (1 + nextMonthGrowthRate);

                    // Add the report date to the list and sheet
                    var newEntry = new NavEvolutionEntry
                    {
                        Date = reportDate,
                        NavPerShare = nextValue
                    };
                    navEvolution.Add(newEntry);

                    // Write to sheet
                    sheet.Cell(row, 1).Value = newEntry.Date;
                    sheet.Cell(row, 2).Value = newEntry.NavPerShare;

                    workbook.Save();

                    Console.WriteLine($"✅ Added NAV for {reportDate:yyyy-MM-dd} -> {nextValue:F2}");
                }
                else
                {
                    Console.WriteLine($"ℹ️ NAV for {reportDate:yyyy-MM-dd} already exists. No changes made.");
                }
            }

            return navEvolution;
        }

        public void UpdateAndSaveHistoricalPerformance(DateTime reportDate, decimal monthlyReturn)
        {
            if (!File.Exists(HistoricalPerformanceFilePath))
            {
                throw new FileNotFoundException($"Historical Performance file not found at: {HistoricalPerformanceFilePath}");
            }

            int colIndex = reportDate.Month + 1;
            int rowIndex = reportDate.Year - 2024 + 2;

            using (var workbook = new XLWorkbook(HistoricalPerformanceFilePath))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new InvalidOperationException("No worksheet found in the Historical Performance file.");
                }

                worksheet.Cell(rowIndex, colIndex).SetValue(monthlyReturn);

                workbook.Save();
            }
        }

        private decimal ExtractDecimalFromCell(IXLWorksheet ws, string cellAddress)
        {
            var cell = ws.Cell(cellAddress);
            if (cell.TryGetValue(out double value))
            {
                return (decimal)value;
            }
            return 0m;
        }

        private List<MonthlyReturn> ExtractMonthlyPerformanceGrid(DateTime reportDate, decimal currentMonthlyReturn)
        {
            var monthlyReturns = new List<MonthlyReturn>();

            if (!File.Exists(HistoricalPerformanceFilePath))
            {
                throw new FileNotFoundException($"Historical Performance file not found for grid extraction: {HistoricalPerformanceFilePath}");
            }

            using (var workbook = new XLWorkbook(HistoricalPerformanceFilePath))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet == null) return monthlyReturns;

                int startCol = 2;
                int endCol = 13;

                var rowsUsed = worksheet.Column(1).CellsUsed().Select(c => c.WorksheetRow().RowNumber()).Where(r => r >= 2);

                foreach (var rowNum in rowsUsed)
                {
                    if (!worksheet.Cell(rowNum, 1).TryGetValue(out int year)) continue;

                    var reportYear = reportDate.Year;
                    var reportMonth = reportDate.Month;

                    for (int monthIndex = 1; monthIndex <= 12; monthIndex++)
                    {
                        // Skip months beyond the report date for the current year
                        if (year == reportYear && monthIndex > reportMonth)
                            break; // or continue; depending on whether you want to skip all months after

                        int colNum = startCol + monthIndex - 1;

                        double returnDouble;
                        if (worksheet.Cell(rowNum, colNum).TryGetValue(out returnDouble))
                        {
                            var date = new DateTime(year, monthIndex, 1);
                            decimal percentageValue = (decimal)returnDouble * 100m;
                            decimal roundedPercentage = Decimal.Round(percentageValue, 2);
                            monthlyReturns.Add(new MonthlyReturn
                            {
                                Date = date,
                                Return = roundedPercentage,
                                MonthName = date.ToString("MMM")
                            });
                        }
                    }
                }
            }

            return monthlyReturns.OrderBy(r => r.Date).ToList();
        }

        private List<NavEvolutionEntry> ExtractNavEvolution(IXLWorksheet ws, string rangeAddress)
        {
            var list = new List<NavEvolutionEntry>();
            var range = ws.Range(rangeAddress);

            foreach (var row in range.Rows())
            {
                if (row.Cell(1).TryGetValue(out DateTime date) && row.Cell(2).TryGetValue(out double navDouble))
                {
                    list.Add(new NavEvolutionEntry
                    {
                        Date = date,
                        NavPerShare = navDouble
                    });
                }
            }
            return list;
        }

        private void ExtractCounterpartyRisks(FactSheetData factData, IXLWorksheet worksheet)
        {
            var counterpartyMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Hercle", "Hercle" },
                { "Citi", "Citi" },
                { "KrakenSpot", "Kraken" },
                { "KrakenCoinFuture", "Kraken" },
                { "KrakenUsdFuture", "Kraken" },
                { "BinanceSpot", "Binance" },
                { "BinanceUSDFuture", "Binance" },
                { "Bitmex", "Bitmex" },
                { "Deribit", "Deribit" },
                { "Bybit", "Bybit" }
            };

            var exposures = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            var lastRow = worksheet.Column("B").LastCellUsed().Address.RowNumber;

            for (int row = 2; row <= lastRow; row++)
            {
                var rawName = worksheet.Cell(row, "B").GetString();
                var normalizedName = counterpartyMapping.ContainsKey(rawName) ? counterpartyMapping[rawName] : rawName;

                decimal exposureValue = 0;
                var cellValue = worksheet.Cell(row, "L").GetValue<string>();

                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    if (!decimal.TryParse(cellValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out exposureValue))
                    {
                        exposureValue = 0;
                    }
                }

                if (exposures.ContainsKey(normalizedName))
                    exposures[normalizedName] += exposureValue;
                else
                    exposures[normalizedName] = exposureValue;
            }

            decimal totalNav = factData.CurrentAuM_USD;

            foreach (var kvp in exposures)
            {
                factData.CounterpartyRisks.Add(new CounterpartyRiskEntry
                {
                    Counterparty = kvp.Key,
                    ExposureUSD = kvp.Value,
                    ExposurePercentNAV = totalNav != 0
                        ? (kvp.Value / totalNav) * 100
                        : 0
                });
            }
            factData.CounterpartyRisks = factData.CounterpartyRisks
                .OrderByDescending(c => c.ExposureUSD)
                .Take(5)
                .ToList();
        }

        private void ExtractAssetBreakdown(FactSheetData factData, IXLWorksheet worksheet)
        {
            var assetValues = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            var lastRow = worksheet.Column("F").LastCellUsed().Address.RowNumber;

            for (int row = 2; row <= lastRow; row++)
            {
                var assetName = worksheet.Cell(row, "F").GetString()?.Trim();
                if (string.IsNullOrWhiteSpace(assetName))
                    continue;

                var cellValue = worksheet.Cell(row, "L").GetValue<string>();
                decimal valueUSD = 0;

                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    if (!decimal.TryParse(cellValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out valueUSD))
                    {
                        valueUSD = 0;
                    }
                }

                if (assetValues.ContainsKey(assetName))
                    assetValues[assetName] += valueUSD;
                else
                    assetValues[assetName] = valueUSD;
            }

            decimal totalNav = factData.CurrentAuM_USD;

            foreach (var kvp in assetValues.OrderByDescending(x => x.Value))
            {
                factData.AssetsBreakdown.Add(new AssetBreakdownEntry
                {
                    Category = kvp.Key,
                    ValueUSD = kvp.Value,
                    ValuePercent = totalNav != 0
                        ? (kvp.Value / totalNav) 
                        : 0
                });
            }
        }

        public void GenerateAssetBreakdownChart(FactSheetData data)
        {
            var filteredAssets = data.AssetsBreakdown.Where(a => Math.Round(a.ValuePercent * 100) >= 1).ToList();

            var excelApp = new Excel.Application();
            excelApp.Visible = false; 

            var workbook = excelApp.Workbooks.Add();
            var sheet = (Excel.Worksheet)workbook.Sheets[1];

            sheet.Cells[1, 1] = "Category";
            sheet.Cells[1, 2] = "Value";
            for (int i = 0; i < filteredAssets.Count; i++)
            {
                sheet.Cells[i + 2, 1] = filteredAssets[i].Category;
                sheet.Cells[i + 2, 2] = filteredAssets[i].ValuePercent;
            }

            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            Excel.ChartObject chartObject = charts.Add(25, 25, 517, 323); // left, top, width, height
            Excel.Chart chart = chartObject.Chart;

            string templatePath = @"C:\Users\antho\Desktop\AssetBreakdownChart.crtx";
            chart.ApplyChartTemplate(templatePath);

            Excel.Range chartRange = sheet.Range[
                sheet.Cells[1, 1],
                sheet.Cells[filteredAssets.Count + 1, 2]
            ];
            chart.SetSourceData(chartRange);

            chart.SeriesCollection(1).DataLabels().Font.Size = 15;
            chart.SeriesCollection(1).DataLabels().Font.Color = 0;
            //chart.SeriesCollection(1).DataLabels().Font.Bold = true;

            Excel.PlotArea plotArea = chart.PlotArea;

            // Shrink the plot area so the pie appears smaller
            plotArea.Width = plotArea.Width * 0.8;   // 70% of current width
            plotArea.Height = plotArea.Height * 0.8;
            plotArea.Left = (chart.ChartArea.Width - plotArea.Width) / 2;
            plotArea.Top = (chart.ChartArea.Height - plotArea.Height) / 2;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = System.IO.Path.Combine(desktopPath, "AssetsBreakdown.png");
            chart.Export(filePath, "PNG", false);

            workbook.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine($"Pie chart saved to: {filePath}");
        }

        public string GetHtmlBody(Google.Apis.Gmail.v1.Data.Message message)
        {
            if (message.Payload == null)
                return string.Empty;

            return GetHtmlPart(message.Payload);

            string GetHtmlPart(Google.Apis.Gmail.v1.Data.MessagePart part)
            {
                if (part.MimeType == "text/html" && part.Body?.Data != null)
                {
                    // Gmail API encodes body in Base64Url
                    var bytes = Convert.FromBase64String(part.Body.Data.Replace('-', '+').Replace('_', '/'));
                    return System.Text.Encoding.UTF8.GetString(bytes);
                }

                if (part.Parts != null && part.Parts.Count > 0)
                {
                    foreach (var subPart in part.Parts)
                    {
                        var html = GetHtmlPart(subPart);
                        if (!string.IsNullOrEmpty(html))
                            return html;
                    }
                }

                return string.Empty;
            }
        }

        public async Task<Dictionary<string, decimal>> GetRiskScenarios(DateTime reportDate)
        {
            string[] Scopes = { GmailService.Scope.GmailReadonly };
            string ApplicationName = "Gmail API .NET Quickstart";

            // Load credentials from your client_secret.json
            UserCredential credential;
            using (var stream = new FileStream("C:\\Users\\antho\\Desktop\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    System.Threading.CancellationToken.None
                );
            }

            // Create Gmail API service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Search for the email (adjust query as needed)
            var request = service.Users.Messages.List("me");
            request.Q = $"subject:'[ProfitLoss] - SigmaNeutral Daily PnL Report For {reportDate.ToString("yyyy-MM-dd")}'";
            request.MaxResults = 1;

            var response = await request.ExecuteAsync();

            if (response.Messages == null || response.Messages.Count == 0)
            {
                Console.WriteLine("No email found.");
                return new Dictionary<string, decimal>();
            }

            var messageId = response.Messages.First().Id;
            var message = await service.Users.Messages.Get("me", messageId).ExecuteAsync();

            string htmlBody = GetHtmlBody(message);

            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlBody);
            var risks = new Dictionary<string, decimal>
            {
                { "USDT Depeg (0.9)", 0 },
                { "USDC Depeg (0.9)", 0 },
                { "STETH Depeg (0.9)", 0 },
                { "BETH Depeg (0.9)", 0 }
            }; ;

            foreach (var key in risks.Keys.ToList())
            {
                var td = doc.DocumentNode.SelectSingleNode($"//td[contains(translate(normalize-space(.), '\u00A0', ' '), '{key}')]");

                decimal value = 0;

                if (td != null)
                {
                    var valueTd = td.SelectSingleNode("following-sibling::td[3]");
                    if (valueTd != null)
                    {
                        var text = valueTd.InnerText.Trim().Replace("%", "").Replace(",", ".");
                        decimal.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
                    }
                }

                risks[key] = value;
            }

            return risks;
        }

        public void GenerateNAVEvolutionChart(FactSheetData data)
        {
            var distinctNav = data.NavEvolution
                .GroupBy(n => new { n.Date.Year, n.Date.Month })
                .Select(g => g.First())
                .ToList();

            var lastEntry = distinctNav.Last();
            distinctNav[distinctNav.Count - 1] = new NavEvolutionEntry
            {
                Date = new DateTime(lastEntry.Date.Year, lastEntry.Date.Month, 1), // Aug-1
                NavPerShare = lastEntry.NavPerShare
            };

            var excelApp = new Excel.Application();
            excelApp.Visible = false;

            var workbook = excelApp.Workbooks.Add();
            var sheet = (Excel.Worksheet)workbook.Sheets[1];

            // Write headers
            sheet.Cells[1, 1] = "Date";
            sheet.Cells[1, 2] = "NAV";

            // Fill data from sorted list, including the projected month
            for (int i = 0; i < distinctNav.Count; i++)
            {
                sheet.Cells[i + 2, 1] = distinctNav[i].Date;
                sheet.Cells[i + 2, 2] = distinctNav[i].NavPerShare;
            }

            // Create chart
            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            Excel.ChartObject chartObject = charts.Add(25, 25, 517, 319); // left, top, width, height
            Excel.Chart chart = chartObject.Chart;

            // Set source data
            Excel.Range chartRange = sheet.Range[
                sheet.Cells[1, 1],
                sheet.Cells[distinctNav.Count + 1, 2]
            ];
            chart.SetSourceData(chartRange);

            // Apply NAV evolution template
            string templatePath = @"C:\Users\antho\Desktop\NAVEvolutionChart.crtx";
            chart.ApplyChartTemplate(templatePath);

            chart.SeriesCollection(1).Format.Line.Weight = 4;

            Excel.Axis xAxis = chart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.MinimumScaleIsAuto = true;
            xAxis.MaximumScaleIsAuto = true;
            xAxis.Crosses = Excel.XlAxisCrosses.xlAxisCrossesAutomatic;

            // Ensure it's treated as date axis
            xAxis.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            xAxis.TickLabels.NumberFormat = "mmm-yy";
            xAxis.TickLabels.Orientation = (Excel.XlTickLabelOrientation)45;
            
            xAxis.TickLabels.Font.Bold = true;
            xAxis.TickLabels.Font.Size = 17;

            Excel.Axis yAxis = chart.Axes(Excel.XlAxisType.xlValue);
            yAxis.MinimumScaleIsAuto = true;
            yAxis.MaximumScaleIsAuto = true;

            // Set major unit to 50
            yAxis.MajorUnit = 50;
            yAxis.TickLabels.Font.Bold = true;
            yAxis.TickLabels.Font.Size = 17;

            // Export chart as PNG
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = System.IO.Path.Combine(desktopPath, "NAVEvolution.png");
            chart.Export(filePath, "PNG", false);

            // Cleanup
            workbook.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine($"NAV Evolution chart saved to: {filePath}");
        }

        public async Task<string> GenerateFactSheetHtmlAsync(FactSheetData data)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string htmlPath = Path.Combine(desktop, "test2.html");

            var options = new LaunchOptions
            {
                Headless = true,
                ExecutablePath = @"H:\Shared drives\Melanion\Tools\C.Common Tools\8.Common Modules\chrome-win\chrome.exe",
                Args = new[] { "--disable-blink-features=AutomationControlled" }
            };

            using var browser = await Puppeteer.LaunchAsync(options);
            using var page = await browser.NewPageAsync();

            // Load your local HTML file
            await page.GoToAsync($"file:///{htmlPath.Replace("\\", "/")}", new NavigationOptions
            {
                WaitUntil = new[] { WaitUntilNavigation.Load }
            });

            // Inject data safely via JS
            await page.EvaluateFunctionAsync(@"(data) => {
                const setText = (id, value) => {
                    const el = document.getElementById(id);
                    if(el) el.textContent = value;
                };

                // --- Fill simple fields ---
                document.querySelectorAll('.title-sub').forEach(e => e.textContent = 'Monthly Report ' + data.reportDate);
                setText('currentAum', '$' + Math.floor(data.currentAum / 1000).toLocaleString() + 'K');

                // --- Risks ---
                for(const [k,v] of Object.entries(data.risks)){
                    setText(k + '-Value', v.toFixed(2) + '%');
                }

                // --- Monthly returns ---
                const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
                const yearlyData = {}; // store compounded value per year

                data.monthly.forEach(m => {
                    const id = `${m.year}-${m.month}`;
                    const cell = document.getElementById(id);
                    if(cell) {
                        const valueText = m.value.toFixed(2) + '%';
                        cell.textContent = valueText;
                        cell.style.fontWeight = '500';

                        // background if non-zero
                        if(m.value !== 0) cell.style.backgroundColor = '#ddebf7';
                        else cell.style.backgroundColor = '';

                        // initialize yearly compounding
                        if(!yearlyData[m.year]) yearlyData[m.year] = 1;

                        // compound monthly return
                        yearlyData[m.year] *= (1 + m.value / 100); // divide by 100 because m.value is in %
                    }
                });

                // --- Populate Yearly cells ---
                for(const [year, compounded] of Object.entries(yearlyData)) {
                    const yearlyCell = document.getElementById(`${year}-Yearly`);
                    if(yearlyCell) {
                        const yearlyReturn = (compounded - 1) * 100; // back to percentage
                        yearlyCell.textContent = yearlyReturn.toFixed(2) + '%';
                        yearlyCell.style.fontWeight = '500';
                        if(yearlyReturn !== 0) yearlyCell.style.backgroundColor = '#ddebf7';
                        else yearlyCell.style.backgroundColor = '';
                    }
                }

                // --- Counterparty risks ---
                const tbody = document.querySelector('.table-counterrisk tbody');
                if(tbody){
                    tbody.innerHTML = '';
                    data.counterparties.forEach(c => {
                        tbody.innerHTML += `
                            <tr>
                                <td><strong>${c.name}</strong></td>
                                <td class='text-center' style='font-weight: 500'>${Math.round(c.usd).toLocaleString()}</td>
                                <td class='text-center' style='font-weight: 500'>${c.perc < 1 ? c.perc.toFixed(1) : Math.round(c.perc)}%</td>
                            </tr>`;
                    });
                }

                // --- Charts ---
                const navImg = document.getElementById('nav-evolution-chart');
                const assetImg = document.getElementById('assets-breakdown-chart');
                if(navImg) navImg.src = 'file:///' + data.navChartPath.replace(/\\\\/g,'/');
                if(assetImg) assetImg.src = 'file:///' + data.assetChartPath.replace(/\\\\/g,'/');
            }", new
            {
                reportDate = data.ReportDate.ToString("dd/MM/yyyy"),
                currentAum = data.CurrentAuM_USD,
                risks = data.Risks.ToDictionary(
                    kv => kv.Key.Split(' ')[0],
                    kv => kv.Value
                ),
                monthly = data.MonthlyPerformanceGrid.Select(m => new
                {
                    year = m.Date.Year,
                    month = m.Date.ToString("MMM"),
                    value = m.Return
                }),
                counterparties = data.CounterpartyRisks.Select(c => new
                {
                    name = c.Counterparty,
                    usd = c.ExposureUSD,
                    perc = c.ExposurePercentNAV
                }),
                navChartPath = Path.Combine(desktop, "NAVEvolution.png"),
                assetChartPath = Path.Combine(desktop, "AssetsBreakdown.png")
            });

            // Get the updated HTML from the page
            string populatedHtml = await page.GetContentAsync();

            // Optionally save to desktop
            string outputPath = Path.Combine(desktop, $"FactSheetPopulated_{data.ReportDate:yyyyMM}.html");
            await File.WriteAllTextAsync(outputPath, populatedHtml);

            Console.WriteLine($"✅ Populated HTML saved: {outputPath}");

            return populatedHtml;
        }


        public void CropImage()
        {
            string inputPath = @"C:\Users\antho\Desktop\AssetsBreakdown.png";
            string outputPath = @"C:\Users\antho\Desktop\AssetsBreakdownCrop.png";
            Bitmap original = new Bitmap(inputPath);

            // Desired height
            int targetHeight = 350;
            int targetWidth = 690; // Set target width if needed

            // Calculate how many pixels to remove from top and bottom
            int cropHeight = targetHeight;
            int cropTop = (original.Height - cropHeight) / 2;

            if (cropTop < 0) cropTop = 0; // in case image is smaller than target

            // Define crop rectangle
            Rectangle cropRect = new Rectangle(0, cropTop, original.Width, cropHeight);

            // Optionally, resize width to 690 if needed
            Bitmap cropped = new Bitmap(cropRect.Width, cropRect.Height);
            using (Graphics g = Graphics.FromImage(cropped))
            {
                g.DrawImage(original,
                            new Rectangle(0, 0, cropped.Width, cropped.Height),
                            cropRect,
                            GraphicsUnit.Pixel);
            }

            // Resize width to 690 if you want exact 690x336
            if (original.Width != targetWidth)
            {
                Bitmap resized = new Bitmap(targetWidth, targetHeight);
                using (Graphics g = Graphics.FromImage(resized))
                {
                    g.DrawImage(cropped, 0, 0, targetWidth, targetHeight);
                }
                cropped.Dispose();
                cropped = resized;
            }

            // Save the final cropped image
            cropped.Save(outputPath);

            // Cleanup
            original.Dispose();
            cropped.Dispose();

            Console.WriteLine("Image cropped and resized to 690x336 successfully!");
        }

    }


    class Program
    {
        #region Daily tasks extraction from Gmail
        public static async Task GetDailyTasks()
        {
            string[] Scopes = { GmailService.Scope.GmailReadonly };
            string ApplicationName = "Gmail API .NET Daily Task Extractor";

            // Load credentials from your client_secret.json
            UserCredential credential;
            using (var stream = new FileStream("C:\\Users\\antho\\Desktop\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    System.Threading.CancellationToken.None
                );
            }

            // Create Gmail API service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Search for emails with subject "Daily Activity Report - "
            var request = service.Users.Messages.List("me");
            request.LabelIds = new string[] { "SENT" }; // look in sent emails
            request.Q = "subject:'Daily Activity Report'";
            request.MaxResults = 200;

            var response = await request.ExecuteAsync();

            if (response.Messages == null || response.Messages.Count == 0)
            {
                Console.WriteLine("No emails found.");
                return;
            }

            var allTasks = new StringBuilder();

            foreach (var msg in response.Messages)
            {
                var message = await service.Users.Messages.Get("me", msg.Id).ExecuteAsync();

                // Get full email body (plain text)
                string body = "";

                if (message.Payload.Parts == null && message.Payload.Body != null)
                {
                    body = Base64UrlDecode(message.Payload.Body.Data);
                }
                else if (message.Payload.Parts != null)
                {
                    foreach (var part in message.Payload.Parts)
                    {
                        if (part.MimeType == "text/plain")
                        {
                            body = Base64UrlDecode(part.Body.Data);
                            break;
                        }
                    }
                }

                // Decode HTML entities if needed
                body = body.Replace("&#39;", "'");

                int index = body.IndexOf("Today's Tasks:", StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    // Take everything after "Today's Tasks:"
                    string tasksText = body.Substring(index + "Today's Tasks:".Length).Trim();

                    allTasks.AppendLine($"Email Date: {UnixTimeStampToDateTime(message.InternalDate ?? 0)}");
                    allTasks.AppendLine(tasksText); // write entire text after Today's Tasks
                    allTasks.AppendLine("--------------------------------------------------");
                }
            }


            // Save to Desktop
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var filePath = Path.Combine(desktopPath, "DailyTasks.txt");
            await File.WriteAllTextAsync(filePath, allTasks.ToString());

            Console.WriteLine($"Tasks saved to {filePath}");
        }
        private static string Base64UrlDecode(string input)
        {
            var output = input.Replace('-', '+').Replace('_', '/');
            switch (output.Length % 4)
            {
                case 2: output += "=="; break;
                case 3: output += "="; break;
            }
            var bytes = Convert.FromBase64String(output);
            return Encoding.UTF8.GetString(bytes);
        }
        private static DateTime UnixTimeStampToDateTime(long unixTimeStamp)
        {
            // Gmail API InternalDate is in milliseconds
            DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddMilliseconds(unixTimeStamp).ToLocalTime();
            return dtDateTime;
        }
        #endregion


        
        static async Task Main(string[] args)
        {
            Console.WriteLine("--- Factsheet Data Extractor (ClosedXML) ---");

            DateTime reportDate = new DateTime(2025, 09, 30);
            decimal jul2025Return = 0.0111m;
            var processor = new ExcelDataProcessor();


            Console.WriteLine($"Attempting to extract data for report date: {reportDate:dd/MM/yyyy}");


            FactSheetData data = processor.ExtractData(reportDate, jul2025Return);

            Console.WriteLine("\nDATA EXTRACTION SUCCESSFUL!");
            Console.WriteLine("------------------------------------------");

            processor.GenerateAssetBreakdownChart(data);
            processor.CropImage();
            processor.GenerateNAVEvolutionChart(data);

            await processor.GenerateFactSheetHtmlAsync(data);

            Console.WriteLine("\nChart done!");

            SelectPdf.HtmlToPdf converter = new SelectPdf.HtmlToPdf();

            SelectPdf.PdfDocument doc = converter.ConvertUrl("C:\\Users\\antho\\Desktop\\FactSheetPopulated_202509.html");
            doc.Save("C:\\Users\\antho\\Desktop\\output.pdf");
            doc.Close();
            Console.WriteLine("PDF created successfully!");
            
        }
    }
}
