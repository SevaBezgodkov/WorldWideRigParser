using System.Net.Sockets;
using HtmlAgilityPack;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

//If you run the program more than once, close previously installed CSV and XLSX files. If files are opened while the program is running, they cannot be overwritten.
namespace RigCountConverter
{
    class Program
    {
        const string penultimateYear = "1976";
        static HttpClient httpClient = new HttpClient();
        
        static async Task Main(string[] args)
        {
            string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl";
            //string excelFileName = @"C:\rig_counts.xlsx";
            //string csvFileName = @"С:\rig_counts.csv";

            string excelFileName = @"rig_counts.xlsx";
            string csvFileName = @"rig_counts.csv";

            var htmlContent = await DownloadHtmlAsync(url);

            var excelFileUrl = ParseExcelFileUrl(htmlContent);

            await DownloadExcelFileAsync(excelFileUrl, excelFileName);

            var rigData = ParseExcelFile(excelFileName);

            ConvertToCsv(rigData, csvFileName);

            Console.WriteLine("CSV file created successfully.");
        }

        static async Task<string> DownloadHtmlAsync(string url)
        {
            try
            {
                var cancellationTokenSource = new CancellationTokenSource();
                var cancellationToken = cancellationTokenSource.Token;

                httpClient.BaseAddress = new Uri(url);
                httpClient.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.32.3");
                httpClient.DefaultRequestHeaders.Add("Accept", "*/*");

                var request = await httpClient.GetStringAsync(url, cancellationToken);

                return request;
            }
            catch (SocketException ex)
            {
                throw new Exception(ex.Message);
            }
        }

        static string ParseExcelFileUrl(string htmlContent)
        {
            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(htmlContent);

            var linkNode = htmlDocument.DocumentNode.SelectSingleNode("//a[contains(normalize-space(), 'Worldwide Rig Counts - Current &amp; Historical Data')]");

            return linkNode.GetAttributeValue("href", "");
        }

        static async Task DownloadExcelFileAsync(string excelFileUrl, string fileName)
        {
            var host = httpClient.BaseAddress?.Host;
            var fileLink = new Uri("https://" + host + excelFileUrl);
            var excelData = await httpClient.GetByteArrayAsync(fileLink);
            File.WriteAllBytes(fileName, excelData);
        }

        static ExcelWorksheet ParseExcelFile(string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage(new FileInfo(fileName));
            var worksheet = package.Workbook.Worksheets[0]; 

            return worksheet;
        }


        static void ConvertToCsv(ExcelWorksheet worksheet, string fileName)
        {
            var writer = new StreamWriter(fileName);

            bool foundMarker = false; 

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                var value = worksheet.Cells[row, 2].Value; 
                if (!foundMarker && value != null && value.ToString() == penultimateYear)
                {
                    foundMarker = true; 
                }

                if (foundMarker )
                {
                    string rowData = "";
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowData += $"{worksheet.Cells[row, col].Value},";
                    }
                    writer.WriteLine(rowData.TrimEnd(','));
                }
            }

            writer.Close(); 
        }
    }
}
