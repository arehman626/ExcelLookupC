using OfficeOpenXml;
using ExcelLookupC.Models;
using ExcelLookupC.Services;

namespace ExcelLookupC.TestData
{
    public static class TestDataGenerator
    {
        static TestDataGenerator()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static async Task CreateSampleFilesAsync(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await CreateLeftSampleFileAsync(Path.Combine(directory, "LeftFile.xlsx"));
            await CreateRightSampleFileAsync(Path.Combine(directory, "RightFile.xlsx"));
        }

        private static async Task CreateLeftSampleFileAsync(string filePath)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Employees");

            // Headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Age";
            worksheet.Cells[1, 4].Value = "City";
            worksheet.Cells[1, 5].Value = "Department";

            // Data
            var data = new object[,]
            {
                { 1, "John Smith", 25, "New York", "IT" },
                { 2, "Jane Doe", 30, "Los Angeles", "HR" },
                { 3, "Bob Johnson", 35, "Chicago", "Finance" },
                { 4, "Alice Brown", 28, "Houston", "IT" },
                { 5, "Charlie Wilson", 32, "Phoenix", "Marketing" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = data[i, j];
                }
            }

            // Format headers
            worksheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;
            worksheet.Cells.AutoFitColumns();

            await package.SaveAsAsync(new FileInfo(filePath));
        }

        private static async Task CreateRightSampleFileAsync(string filePath)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Staff");

            // Headers (slightly different names to test column matching)
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Age";
            worksheet.Cells[1, 4].Value = "Location";
            worksheet.Cells[1, 5].Value = "Salary";

            // Data with some matches and differences
            var data = new object[,]
            {
                { 1, "John Smith", 25, "New York", 65000 },
                { 2, "Jane Doe", 31, "Los Angeles", 70000 }, // Age different
                { 4, "Alice Brown", 28, "Houston", 62000 },
                { 6, "David Lee", 29, "Miami", 58000 }, // New person
                { 7, "Emma Davis", 33, "Seattle", 75000 } // New person
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = data[i, j];
                }
            }

            // Format headers
            worksheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;
            worksheet.Cells.AutoFitColumns();

            await package.SaveAsAsync(new FileInfo(filePath));
        }
    }
}