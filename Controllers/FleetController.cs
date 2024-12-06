using System.Data.Common;
using System.Data.Odbc;
using System.Text;

using Microsoft.AspNetCore.Mvc;

using Newtonsoft.Json;

using OfficeOpenXml;

namespace ToolForOtis.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FleetController : ControllerBase
    {
        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file, [FromForm] string username, [FromForm] string password)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Please upload a valid Excel file.");

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
                return BadRequest("Username and password are required.");

            // Parse Excel File
            List<ExcelColumns> uploadedExcelData = new();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                        return BadRequest("Invalid Excel file.");

                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowCount; row++) // Assuming first row is header
                    {
                        string timestampUTC;
                        var cellValue = worksheet.Cells[row, 3].Value;
                        if (cellValue is double numericValue)
                        {
                            DateTime timestamp = DateTime.FromOADate(numericValue);
                            timestampUTC = timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            timestampUTC = worksheet.Cells[row, 3].Text;
                        }

                        uploadedExcelData.Add(
                            new ExcelColumns()
                            {
                                IMO = worksheet.Cells[row, 1].Text,
                                VesselName = worksheet.Cells[row, 2].Text,
                                TimestampUTC = timestampUTC,
                                Report = worksheet.Cells[row, 4].Text
                            });
                    }
                }
            }

            // Step 1: Call the first API to get the fleet data
            var fleetResponse = await CallFleetApi(username, password);
            if (fleetResponse == null || fleetResponse.Data?.Fleet == null)
                return StatusCode(500, "Failed to fetch fleet data.");

            var fleetList = fleetResponse.Data.Fleet;

            // Match IMOs with fleet list
            var validImos = uploadedExcelData
                .Where(x => fleetList.Any(y => x.IMO == y.Imo))
                .Select(x => x.IMO)
                .ToList();

            if (!validImos.Any())
                return NotFound("No matching IMOs found.");

            // Step 2: Match IMOs from Excel with the fleet list to get serials
            var matchedSerials = fleetList
                .Where(x => validImos.Contains(x.Imo))
                .Select(f => f.Serial)
                .ToList();

            if (!matchedSerials.Any())
                return NotFound("No matching serials found for the provided IMOs.");

            // Step 3: Call the second API to get the latest reports
            var reports = await CallReportsApi(username, password, matchedSerials);
            if (reports == null)
                return StatusCode(500, "Failed to fetch report data.");

            foreach (var report in reports)
            {
                var extendedReport = await CallVesselExtendedInformationApi(username, password, report.Serial);
                report.MMSI = extendedReport.Mmsi;
                report.S3LatestTimestamp = (await CallS3AisHistoryApi(extendedReport.Mmsi, report.Timestamp, DateTime.UtcNow.ToString("yyyy-MM-dd")))
                                            .OrderByDescending(r => r.Timestamp).FirstOrDefault()
                                            .Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                report.RedshiftLatestTimestamp = await FetchRedshiftRecord(extendedReport.Mmsi, report.Timestamp);
            }

            var outputFileName = ExcelGenerator.GenerateReportExcel(uploadedExcelData, reports, file.FileName);

            return Ok(new { Message = "Report generated successfully.", FileName = outputFileName });
        }

        private async Task<FleetApiResponse> CallFleetApi(string username, string password)
        {
            using var client = new HttpClient();
            var requestData = $"<otisrequest><login>{username}</login><password>{password}</password><action>GetUserProfile</action><responseformat>json</responseformat><parameters><useragent>Mozilla/5.0</useragent></parameters></otisrequest>";

            var content = new StringContent($"data={requestData}", Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await client.PostAsync("https://otis.stratumfive.com/api/otis.ashx", content);

            if (!response.IsSuccessStatusCode)
                return null;

            var responseString = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<FleetApiResponse>(responseString);
        }

        private async Task<List<ReportResponse>> CallReportsApi(string username, string password, List<string> serials)
        {
            using var client = new HttpClient();
            var serialParams = string.Join("", serials.Select(s => $"<serial>{s}</serial>"));
            var requestData = $"<otisrequest><login>{username}</login><password>{password}</password><action>GetLatestReports</action><responseformat>json</responseformat><parameters>{serialParams}</parameters></otisrequest>";

            var content = new StringContent($"data={requestData}", Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await client.PostAsync("https://otis.stratumfive.com/api/otis.ashx", content);

            if (!response.IsSuccessStatusCode)
                return null;

            var responseString = await response.Content.ReadAsStringAsync();
            var reportsResponse = JsonConvert.DeserializeObject<ReportsApiResponse>(responseString);
            return reportsResponse?.Data?.Reports;
        }

        private async Task<VesselExtendedInformationResponse> CallVesselExtendedInformationApi(string username, string password, string serial)
        {
            using var client = new HttpClient();
            var requestData = $"<otisrequest><login>{username}</login><password>{password}</password><action>GetVesselExtendedInformation</action><responseformat>json</responseformat><parameters><serial>{serial}</serial></parameters></otisrequest>";

            var content = new StringContent($"data={requestData}", Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await client.PostAsync("https://otis.stratumfive.com/api/otis.ashx", content);

            if (!response.IsSuccessStatusCode)
                return null;

            var responseString = await response.Content.ReadAsStringAsync();
            var reportsResponse = JsonConvert.DeserializeObject<VesselExtendedInformationApiResponse>(responseString);
            return reportsResponse?.Data?.Fleet.First();
        }

        private async Task<List<S3AisHistoryApiResponse>> CallS3AisHistoryApi(string mmsi, string fromDate, string toDate)
        {
            using var client = new HttpClient();

            var requestData = $"{{ \"dType\": \"mmsi\", \"includePositions\": \"true\", \"includeStaticAndVoyage\": \"true\", \"ids\": [ {mmsi} ], \"from\": \"{fromDate}\", \"to\": \"{toDate}\" }}";
            var content = new StringContent($"data={requestData}", null, "application/json");
            var response = await client.PostAsync("https://otis.stratumfive.com/api/otis.ashx", content);

            if (!response.IsSuccessStatusCode)
                return null;

            var responseString = await response.Content.ReadAsStringAsync();
            var reportsResponse = JsonConvert.DeserializeObject<List<S3AisHistoryApiResponse>>(responseString);
            return reportsResponse;
        }

        private async Task<string> FetchRedshiftRecord(string mmsi, string fromDate)
        {
            string query = "SELECT reporttimestamp FROM \"redshift-db-prod\".\"public\".\"ais_positions_raw\" " +
                $"WHERE mmsi = '{mmsi}' " +
                $"AND reporttimestamp >= '{fromDate}' " +
                "ORDER By reporttimestamp DESC LIMIT 1;";

            using (OdbcConnection connection = new OdbcConnection("Driver=Amazon Redshift ODBC Driver (x64); Server=redshift-cluster-prod.cxmn8ifmencz.us-east-1.redshift.amazonaws.com; Database=redshift-db-prod;UID=redshiftadmin;PWD=4WXKecwm5CRIHHs;Port=5439;SSL=true;"))
            {
                connection.Open();
                Console.WriteLine("Connection Successful!");

                // Execute the query
                using (OdbcCommand command = new OdbcCommand(query, connection))
                {
                    using (OdbcDataReader reader = command.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            return reader[1] as string;
                        }
                    }
                }
            }

            return string.Empty;
        }
    }

    // Models
    public class FleetApiResponse
    {
        public FleetData Data { get; set; }
    }

    public class FleetData
    {
        public List<FleetItem> Fleet { get; set; }
    }

    public class FleetItem
    {
        public string Serial { get; set; }
        public string Imo { get; set; }
    }

    public class ExcelColumns
    {
        public string IMO { get; set; }
        public string VesselName { get; set; }
        public string TimestampUTC {  get; set; }
        public string Report {  get; set; }
    }

    public class ReportsApiResponse
    {
        public ReportsData Data { get; set; }
    }

    public class ReportsData
    {
        public List<ReportResponse> Reports { get; set; }
    }

    public class ReportResponse
    {
        public string ReportId { get; set; }
        public string Serial { get; set; }
        public string Imo { get; set; }
        public string Name { get; set; }
        public string Timestamp { get; set; }
        public string Description { get; set; }
        public string MMSI { get; set; }
        public string RedshiftLatestTimestamp { get; set; }
        public string S3LatestTimestamp { get; set; }
    }

    public class VesselExtendedInformationApiResponse
    {
        public VesselExtendedInformationData Data { get; set; }
    }

    public class VesselExtendedInformationData
    {
        public List<VesselExtendedInformationResponse> Fleet {  get; set; }
    }

    public class VesselExtendedInformationResponse 
    {
        public string Serial { get; set; }
        public string Name { get; set; }
        public object Division { get; set; }
        public string Imo { get; set; }
        public string Mmsi { get; set; }
    }

    public class S3AisHistoryApiResponse
    {
        public string Mmsi { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
