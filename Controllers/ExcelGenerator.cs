using OfficeOpenXml;

namespace ToolForOtis.Controllers
{
    public class ExcelGenerator
    {
        public static string GenerateReportExcel(List<ExcelColumns> excelColumnValues, List<ReportResponse> reports, string uploadedFileName)
        {
            // Create the output file name based on the uploaded file
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(uploadedFileName);
            string outputFileName = $"{fileNameWithoutExtension}_Report.xlsx";

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Reports");

            // Add headers
            string[] headers = { "IMO", "VesselName", "Timestamp", "Report", "Checked At UTC Timestamp", "Last Data Received UTC Timestamp", "Gap from Last Check Time", "Response Report" };
            for (int i = 0; i < headers.Length; i++)
                worksheet.Cells[1, i + 1].Value = headers[i];

            // Add data
            for (int i = 0; i < excelColumnValues.Count; i++)
            {
                var item = excelColumnValues[i];

                if(string.IsNullOrEmpty(item.IMO))
                {
                    continue;
                }

                var report = reports.FirstOrDefault(r => r.Imo == item.IMO && r.Description.Equals(item.Report, StringComparison.OrdinalIgnoreCase));

                worksheet.Cells[i + 2, 1].Value = item.IMO;
                worksheet.Cells[i + 2, 2].Value = item.VesselName;
                worksheet.Cells[i + 2, 3].Value = item.TimestampUTC;
                worksheet.Cells[i + 2, 4].Value = item.Report;
                worksheet.Cells[i + 2, 5].Value = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
                worksheet.Cells[i + 2, 6].Value = report?.Timestamp;
                worksheet.Cells[i + 2, 7].Value = report != null ? (DateTime.UtcNow - DateTime.Parse(report.Timestamp)).ToString() : "N/A";
                worksheet.Cells[i + 2, 8].Value = report?.Description;
            }

            string outputDirectory = Directory.GetCurrentDirectory() + "\\Reports";
            Directory.CreateDirectory(outputDirectory);

            // Save the Excel file
            string outputPath = Path.Combine(outputDirectory, outputFileName);
            
            package.SaveAs(new FileInfo(outputPath));

            return outputFileName;
        }
    }
}
