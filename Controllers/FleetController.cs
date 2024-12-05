using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
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
            List<string> imoList = new();
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
                        imoList.Add(worksheet.Cells[row, 1].Text);
                    }
                }
            }

            // Step 1: Call the first API to get the fleet data
            var fleetResponse = await CallFleetApi(username, password);
            if (fleetResponse == null || fleetResponse.Data?.Fleet == null)
                return StatusCode(500, "Failed to fetch fleet data.");

            var fleetList = fleetResponse.Data.Fleet;

            // Step 2: Match IMOs from Excel with the fleet list to get serials
            var matchedSerials = fleetList
                .Where(f => imoList.Contains(f.Imo))
                .Select(f => f.Serial)
                .ToList();

            if (!matchedSerials.Any())
                return NotFound("No matching serials found for the provided IMOs.");

            // Step 3: Call the second API to get the latest reports
            var reports = await CallReportsApi(username, password, matchedSerials);
            if (reports == null)
                return StatusCode(500, "Failed to fetch report data.");

            return Ok(reports);
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

        private async Task<List<Report>> CallReportsApi(string username, string password, List<string> serials)
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

    public class ReportsApiResponse
    {
        public ReportsData Data { get; set; }
    }

    public class ReportsData
    {
        public List<Report> Reports { get; set; }
    }

    public class Report
    {
        public string ReportId { get; set; }
        public string Serial { get; set; }
        public string Imo { get; set; }
        public string Name { get; set; }
        public string Timestamp { get; set; }
        public string Description { get; set; }
    }
}
