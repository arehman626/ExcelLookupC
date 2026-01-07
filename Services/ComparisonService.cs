using ExcelLookupC.Models;

namespace ExcelLookupC.Services
{
    public class ComparisonService
    {
        public ComparisonResult CompareSheets(SheetData leftSheet, SheetData rightSheet, 
            List<string> comparisonColumns, string leftFileName, string rightFileName)
        {
            var result = new ComparisonResult
            {
                LeftFileName = leftFileName,
                RightFileName = rightFileName,
                LeftSheetName = leftSheet.SheetName,
                RightSheetName = rightSheet.SheetName,
                ComparisonColumns = comparisonColumns,
                ComparisonTimestamp = DateTime.Now,
                LeftSheetData = leftSheet,
                RightSheetData = rightSheet
            };

            // Validate comparison columns exist in both sheets
            var missingInLeft = comparisonColumns.Where(col => !leftSheet.ColumnNames.Contains(col)).ToList();
            var missingInRight = comparisonColumns.Where(col => !rightSheet.ColumnNames.Contains(col)).ToList();

            if (missingInLeft.Any() || missingInRight.Any())
            {
                throw new ArgumentException($"Comparison columns missing. Left: [{string.Join(", ", missingInLeft)}], Right: [{string.Join(", ", missingInRight)}]");
            }

            // Create lookup dictionaries based on comparison columns
            var leftLookup = CreateLookupDictionary(leftSheet.Records, comparisonColumns);
            var rightLookup = CreateLookupDictionary(rightSheet.Records, comparisonColumns);

            // Find matches
            foreach (var leftEntry in leftLookup)
            {
                if (rightLookup.ContainsKey(leftEntry.Key))
                {
                    // Match found
                    var leftRecord = leftEntry.Value;
                    var rightRecord = rightLookup[leftEntry.Key];
                    
                    // Combine records with prefixes for clarity
                    var combinedRecord = new Dictionary<string, object?>();
                    
                    foreach (var kvp in leftRecord)
                    {
                        combinedRecord[$"Left_{kvp.Key}"] = kvp.Value;
                    }
                    
                    foreach (var kvp in rightRecord)
                    {
                        combinedRecord[$"Right_{kvp.Key}"] = kvp.Value;
                    }
                    
                    result.MatchedRecords.Add(combinedRecord);
                    rightLookup.Remove(leftEntry.Key);
                }
                else
                {
                    // Left only
                    result.LeftOnlyRecords.Add(leftEntry.Value);
                }
            }

            // Remaining items in rightLookup are right only
            result.RightOnlyRecords.AddRange(rightLookup.Values);

            // Sort all result lists by comparison columns
            result.MatchedRecords = SortRecords(result.MatchedRecords, comparisonColumns, "Left_");
            result.LeftOnlyRecords = SortRecords(result.LeftOnlyRecords, comparisonColumns);
            result.RightOnlyRecords = SortRecords(result.RightOnlyRecords, comparisonColumns);

            return result;
        }

        public ComparisonResult CompareSheetsWithMapping(SheetData leftSheet, SheetData rightSheet, 
            List<ColumnMapping> columnMappings, string leftFileName, string rightFileName)
        {
            var result = new ComparisonResult
            {
                LeftFileName = leftFileName,
                RightFileName = rightFileName,
                LeftSheetName = leftSheet.SheetName,
                RightSheetName = rightSheet.SheetName,
                ColumnMappings = columnMappings,
                ComparisonColumns = columnMappings.Select(m => m.LeftColumn).ToList(),
                ComparisonTimestamp = DateTime.Now,
                LeftSheetData = leftSheet,
                RightSheetData = rightSheet
            };

            // Validate mapping columns exist in their respective sheets
            var missingInLeft = columnMappings.Where(m => !leftSheet.ColumnNames.Contains(m.LeftColumn)).ToList();
            var missingInRight = columnMappings.Where(m => !rightSheet.ColumnNames.Contains(m.RightColumn)).ToList();

            if (missingInLeft.Any() || missingInRight.Any())
            {
                var leftMissing = string.Join(", ", missingInLeft.Select(m => m.LeftColumn));
                var rightMissing = string.Join(", ", missingInRight.Select(m => m.RightColumn));
                throw new ArgumentException($"Mapped columns missing. Left: [{leftMissing}], Right: [{rightMissing}]");
            }

            // Create lookup dictionaries based on mapped columns
            var leftLookup = CreateMappedLookupDictionary(leftSheet.Records, columnMappings, true);
            var rightLookup = CreateMappedLookupDictionary(rightSheet.Records, columnMappings, false);

            // Find matches
            foreach (var leftEntry in leftLookup)
            {
                if (rightLookup.ContainsKey(leftEntry.Key))
                {
                    // Match found
                    var leftRecord = leftEntry.Value;
                    var rightRecord = rightLookup[leftEntry.Key];
                    
                    // Combine records with prefixes for clarity
                    var combinedRecord = new Dictionary<string, object?>();
                    
                    foreach (var kvp in leftRecord)
                    {
                        combinedRecord[$"Left_{kvp.Key}"] = kvp.Value;
                    }
                    
                    foreach (var kvp in rightRecord)
                    {
                        combinedRecord[$"Right_{kvp.Key}"] = kvp.Value;
                    }
                    
                    result.MatchedRecords.Add(combinedRecord);
                    rightLookup.Remove(leftEntry.Key);
                }
                else
                {
                    // Left only
                    result.LeftOnlyRecords.Add(leftEntry.Value);
                }
            }

            // Remaining items in rightLookup are right only
            result.RightOnlyRecords.AddRange(rightLookup.Values);

            // Sort all result lists by mapped columns
            var leftColumns = columnMappings.Select(m => m.LeftColumn).ToList();
            result.MatchedRecords = SortRecords(result.MatchedRecords, leftColumns, "Left_");
            result.LeftOnlyRecords = SortRecords(result.LeftOnlyRecords, leftColumns);
            
            var rightColumns = columnMappings.Select(m => m.RightColumn).ToList();
            result.RightOnlyRecords = SortRecords(result.RightOnlyRecords, rightColumns);

            return result;
        }

        private Dictionary<string, Dictionary<string, object?>> CreateLookupDictionary(
            List<Dictionary<string, object?>> records, List<string> keyColumns)
        {
            var lookup = new Dictionary<string, Dictionary<string, object?>>();

            foreach (var record in records)
            {
                var keyParts = new List<string>();
                foreach (var column in keyColumns)
                {
                    var value = record.ContainsKey(column) ? record[column]?.ToString() ?? "" : "";
                    keyParts.Add(value);
                }

                var compositeKey = string.Join("|", keyParts);
                
                // Handle duplicate keys by appending a counter
                var originalKey = compositeKey;
                int counter = 1;
                while (lookup.ContainsKey(compositeKey))
                {
                    compositeKey = $"{originalKey}_{counter}";
                    counter++;
                }

                lookup[compositeKey] = record;
            }

            return lookup;
        }

        private Dictionary<string, Dictionary<string, object?>> CreateMappedLookupDictionary(
            List<Dictionary<string, object?>> records, List<ColumnMapping> columnMappings, bool useLeftColumns)
        {
            var lookup = new Dictionary<string, Dictionary<string, object?>>();

            foreach (var record in records)
            {
                var keyParts = new List<string>();
                foreach (var mapping in columnMappings)
                {
                    var column = useLeftColumns ? mapping.LeftColumn : mapping.RightColumn;
                    var value = record.ContainsKey(column) ? record[column]?.ToString() ?? "" : "";
                    keyParts.Add(value);
                }

                var compositeKey = string.Join("|", keyParts);
                
                // Handle duplicate keys by appending a counter
                var originalKey = compositeKey;
                int counter = 1;
                while (lookup.ContainsKey(compositeKey))
                {
                    compositeKey = $"{originalKey}_{counter}";
                    counter++;
                }

                lookup[compositeKey] = record;
            }

            return lookup;
        }

        private List<Dictionary<string, object?>> SortRecords(List<Dictionary<string, object?>> records, 
            List<string> sortColumns, string prefix = "")
        {
            if (!records.Any() || !sortColumns.Any())
                return records;

            return records.OrderBy(record =>
            {
                var sortValues = new List<string>();
                foreach (var column in sortColumns)
                {
                    var columnKey = prefix + column;
                    if (record.ContainsKey(columnKey))
                    {
                        sortValues.Add(record[columnKey]?.ToString() ?? "");
                    }
                    else if (record.ContainsKey(column))
                    {
                        sortValues.Add(record[column]?.ToString() ?? "");
                    }
                    else
                    {
                        sortValues.Add("");
                    }
                }
                return string.Join("|", sortValues);
            }).ToList();
        }
    }
}