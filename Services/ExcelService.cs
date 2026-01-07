using OfficeOpenXml;
using ExcelLookupC.Models;
using System.Text;

namespace ExcelLookupC.Services
{
    public class ExcelService
    {
        static ExcelService()
        {
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<ExcelFileInfo> LoadExcelFileAsync(string filePath)
        {
            try
            {
                var fileInfo = new ExcelFileInfo
                {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath)
                };

                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".csv")
                {
                    await LoadCsvFileInfoAsync(filePath, fileInfo);
                }
                else
                {
                    await Task.Run(() =>
                    {
                        using var package = new ExcelPackage(new FileInfo(filePath));
                    
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            fileInfo.SheetNames.Add(worksheet.Name);
                            
                            var columns = new List<string>();
                            if (worksheet.Dimension != null)
                            {
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                                    columns.Add(headerValue);
                                }
                            }
                            fileInfo.SheetColumns[worksheet.Name] = columns;
                        }
                    });
                }

                return fileInfo;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading file: {ex.Message}", ex);
            }
        }

        public async Task<SheetData> LoadSheetDataAsync(string filePath, string sheetName)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".csv")
                {
                    return await LoadCsvSheetDataAsync(filePath, sheetName);
                }
                else
                {
                    return await Task.Run(() =>
                    {
                        var sheetData = new SheetData { SheetName = sheetName };
                        
                        using var package = new ExcelPackage(new FileInfo(filePath));
                        var worksheet = package.Workbook.Worksheets[sheetName];
                    
                        if (worksheet?.Dimension == null)
                        {
                            return sheetData;
                        }

                        // Get column names from first row
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                            sheetData.ColumnNames.Add(headerValue);
                        }

                        // Get data rows
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var record = new Dictionary<string, object?>();
                            bool hasData = false;

                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                var columnName = sheetData.ColumnNames[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value;
                                record[columnName] = cellValue;
                            
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    hasData = true;
                                }
                            }

                            if (hasData)
                            {
                                sheetData.Records.Add(record);
                            }
                        }

                        return sheetData;
                    });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading sheet data: {ex.Message}", ex);
            }
        }

        public async Task ExportComparisonResultAsync(ComparisonResult result, string outputPath)
        {
            try
            {
                using var package = new ExcelPackage();

                // Create summary sheet
                var summarySheet = package.Workbook.Worksheets.Add("Summary");
                summarySheet.Cells[1, 1].Value = "Comparison Summary";
                summarySheet.Cells[2, 1].Value = "Left File:";
                summarySheet.Cells[2, 2].Value = result.LeftFileName;
                summarySheet.Cells[3, 1].Value = "Right File:";
                summarySheet.Cells[3, 2].Value = result.RightFileName;
                summarySheet.Cells[4, 1].Value = "Left Sheet:";
                summarySheet.Cells[4, 2].Value = result.LeftSheetName;
                summarySheet.Cells[5, 1].Value = "Right Sheet:";
                summarySheet.Cells[5, 2].Value = result.RightSheetName;
                summarySheet.Cells[6, 1].Value = "Comparison Columns:";
                summarySheet.Cells[6, 2].Value = string.Join(", ", result.ComparisonColumns);
                summarySheet.Cells[7, 1].Value = "Comparison Time:";
                summarySheet.Cells[7, 2].Value = result.ComparisonTimestamp.ToString("yyyy-MM-dd HH:mm:ss");
                summarySheet.Cells[9, 1].Value = "Matched Records:";
                summarySheet.Cells[9, 2].Value = result.MatchedRecords.Count;
                summarySheet.Cells[10, 1].Value = "Left Only Records:";
                summarySheet.Cells[10, 2].Value = result.LeftOnlyRecords.Count;
                summarySheet.Cells[11, 1].Value = "Right Only Records:";
                summarySheet.Cells[11, 2].Value = result.RightOnlyRecords.Count;

                // Format summary sheet
                summarySheet.Cells[1, 1].Style.Font.Bold = true;
                summarySheet.Cells[1, 1].Style.Font.Size = 16;
                summarySheet.Cells["A2:A11"].Style.Font.Bold = true;
                summarySheet.Cells.AutoFitColumns();

                // Create matched records sheet
                if (result.MatchedRecords.Any())
                {
                    CreateDataSheet(package, "Matched Records", result.MatchedRecords);
                }

                // Create left only records sheet
                if (result.LeftOnlyRecords.Any())
                {
                    CreateDataSheet(package, "Left Only Records", result.LeftOnlyRecords);
                }

                // Create right only records sheet
                if (result.RightOnlyRecords.Any())
                {
                    CreateDataSheet(package, "Right Only Records", result.RightOnlyRecords);
                }

                // Create original left sheet data
                if (result.LeftSheetData != null && result.LeftSheetData.Records.Any())
                {
                    CreateDataSheet(package, "Original Left Data", result.LeftSheetData.Records);
                }

                // Create original right sheet data
                if (result.RightSheetData != null && result.RightSheetData.Records.Any())
                {
                    CreateDataSheet(package, "Original Right Data", result.RightSheetData.Records);
                }

                await package.SaveAsAsync(new FileInfo(outputPath));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error exporting results: {ex.Message}", ex);
            }
        }

        private void CreateDataSheet(ExcelPackage package, string sheetName, List<Dictionary<string, object?>> data)
        {
            var worksheet = package.Workbook.Worksheets.Add(sheetName);
            
            if (!data.Any()) return;

            // Get all unique column names
            var allColumns = data.SelectMany(r => r.Keys).Distinct().OrderBy(k => k).ToList();

            // Add headers
            for (int col = 0; col < allColumns.Count; col++)
            {
                worksheet.Cells[1, col + 1].Value = allColumns[col];
                worksheet.Cells[1, col + 1].Style.Font.Bold = true;
            }

            // Add data
            for (int row = 0; row < data.Count; row++)
            {
                var record = data[row];
                for (int col = 0; col < allColumns.Count; col++)
                {
                    var columnName = allColumns[col];
                    if (record.ContainsKey(columnName))
                    {
                        worksheet.Cells[row + 2, col + 1].Value = record[columnName];
                    }
                }
            }

            worksheet.Cells.AutoFitColumns();
        }

        private async Task LoadCsvFileInfoAsync(string filePath, ExcelFileInfo fileInfo)
        {
            await Task.Run(() =>
            {
                var sheetName = Path.GetFileNameWithoutExtension(filePath);
                fileInfo.SheetNames.Add(sheetName);

                var columns = new List<string>();
                using var reader = new StreamReader(filePath);
                var headerLine = reader.ReadLine();
                
                if (!string.IsNullOrEmpty(headerLine))
                {
                    columns = ParseCsvLine(headerLine);
                }

                fileInfo.SheetColumns[sheetName] = columns;
            });
        }

        private async Task<SheetData> LoadCsvSheetDataAsync(string filePath, string sheetName)
        {
            return await Task.Run(() =>
            {
                var sheetData = new SheetData { SheetName = sheetName };
                
                using var reader = new StreamReader(filePath);
                var headerLine = reader.ReadLine();
                
                if (string.IsNullOrEmpty(headerLine))
                {
                    return sheetData;
                }

                sheetData.ColumnNames = ParseCsvLine(headerLine);

                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var values = ParseCsvLine(line);
                    var record = new Dictionary<string, object?>();
                    bool hasData = false;

                    for (int i = 0; i < sheetData.ColumnNames.Count; i++)
                    {
                        var columnName = sheetData.ColumnNames[i];
                        var value = i < values.Count ? values[i] : "";
                        
                        record[columnName] = string.IsNullOrEmpty(value) ? null : value;
                        
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            hasData = true;
                        }
                    }

                    if (hasData)
                    {
                        sheetData.Records.Add(record);
                    }
                }

                return sheetData;
            });
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];
                
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // Escaped quote
                        current.Append('"');
                        i++; // Skip next quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
            
            result.Add(current.ToString());
            return result;
        }
    }
}