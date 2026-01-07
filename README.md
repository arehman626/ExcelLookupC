# Excel Lookup - File Comparison Tool

A powerful C# Windows Forms application for comparing Excel files with an intuitive dual-panel interface and comprehensive result export capabilities.

## Features

- **Dual Panel Interface**: Select and compare two Excel files side by side
- **Sheet Selection**: Choose which sheets to compare from each file
- **Column Selection**: Select specific columns for comparison using checkboxes
- **Intelligent Comparison**: 
  - Finds matched records between files
  - Identifies records that exist only in the left file
  - Identifies records that exist only in the right file
- **Sorted Results**: All comparison results are automatically sorted for easy analysis
- **Comprehensive Export**: Export results to a timestamped Excel file with:
  - Summary sheet with comparison details
  - Separate sheets for matched, left-only, and right-only records
  - Original file names and metadata preserved

## Requirements

- Windows Operating System
- .NET 8.0 Runtime
- Microsoft Office Excel or compatible application (for viewing results)

## Installation

### Option 1: Clone and Build from Source
1. Clone the repository:
   ```bash
   git clone https://github.com/arehman626/ExcelLookupC.git
   cd ExcelLookupC
   ```

2. Build the application:
   ```bash
   dotnet restore
   dotnet build --configuration Release
   ```

3. Run the application:
   ```bash
   dotnet run --configuration Release
   ```

### Option 2: Download Release Binary
1. Download the latest release from the GitHub releases page
2. Choose your preferred version:
   - **Single File**: `single-file/ExcelLookupC.exe` (158MB) - Single executable, no dependencies
   - **Standalone**: `standalone/` folder - Self-contained with separate files
   - **Framework Dependent**: `publish/` folder - Requires .NET 8.0 runtime
3. Extract (if needed) and run `ExcelLookupC.exe`

### Option 3: Direct Download
- **Single Executable**: Download `single-file/ExcelLookupC.exe` - runs immediately, no installation required
- **Best for**: Quick testing, simple distribution, systems without .NET

## Usage

### Step 1: Load Excel Files
1. **Left Panel**: Click "Select File" to choose your first Excel file
2. **Right Panel**: Click "Select File" to choose your second Excel file
3. Wait for the files to load (you'll see sheet names populated)

### Step 2: Choose Sheets
1. Select the sheet you want to compare from the **Left File** dropdown
2. Select the sheet you want to compare from the **Right File** dropdown
3. The application will automatically identify common columns between the sheets

### Step 3: Select Comparison Columns
1. In the **Comparison Settings** panel, check the columns you want to use for comparison
2. You can select multiple columns - the comparison will be based on the combination of all selected columns
3. Only columns that exist in both sheets will be available for selection

### Step 4: Run Comparison
1. Click the **Compare** button to start the comparison process
2. The progress bar will show the operation status
3. Results will be displayed in the **Comparison Results** panel:
   - **Matched Records**: Records found in both files
   - **Left Only Records**: Records found only in the left file
   - **Right Only Records**: Records found only in the right file

### Step 5: Export Results
1. Click **Export Results** to save the comparison results
2. Choose your desired save location and filename
3. The exported Excel file will contain:
   - **Summary**: Overview of the comparison with metadata
   - **Matched Records**: All matching records with data from both files
   - **Left Only Records**: Records unique to the left file
   - **Right Only Records**: Records unique to the right file

## File Format Support

- **Input**: Excel files (.xlsx, .xls)
- **Output**: Excel files (.xlsx)

## Technical Details

### Architecture
- **Frontend**: Windows Forms (.NET 8.0)
- **Excel Processing**: EPPlus library
- **Comparison Engine**: Custom algorithm with sorted output
- **Export Engine**: Multi-sheet Excel generation with formatting

### Performance
- Asynchronous file operations to prevent UI freezing
- Efficient memory usage for large datasets
- Progress indication for long-running operations

### Error Handling
- Comprehensive error messages for file access issues
- Validation of sheet and column selections
- Graceful handling of malformed Excel files

## Sample Files for Testing

Create test Excel files with the following structure:

### Sample File 1 (LeftFile.xlsx)
| ID | Name | Age | City |
|----|------|-----|------|
| 1 | John | 25 | NYC |
| 2 | Jane | 30 | LA |
| 3 | Bob | 35 | Chicago |

### Sample File 2 (RightFile.xlsx)
| ID | Name | Age | Location |
|----|------|-----|----------|
| 1 | John | 25 | NYC |
| 2 | Jane | 31 | LA |
| 4 | Alice | 28 | Boston |

**Expected Results when comparing on ID + Name:**
- **Matched**: 2 records (John, Jane)
- **Left Only**: 1 record (Bob)
- **Right Only**: 1 record (Alice)

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Troubleshooting

### Common Issues

**"File is locked or in use"**
- Close the Excel file in any other applications before loading
- Ensure you have read permissions on the file

**"No common columns found"**
- Verify that the selected sheets have matching column headers
- Column names are case-insensitive but must match exactly otherwise

**"Application won't start"**
- Ensure .NET 8.0 Runtime is installed
- Check that Windows Forms is supported on your system

### Support

For issues and questions:
1. Check the [Issues](https://github.com/arehman626/ExcelLookupC/issues) page
2. Create a new issue with:
   - Error message (if any)
   - Steps to reproduce
   - Sample files (if possible)

## Version History

### v1.0.0 (Initial Release)
- Dual panel Excel file comparison
- Sheet and column selection
- Sorted comparison results
- Timestamped export functionality
- Comprehensive error handling
- Windows Forms GUI

---

**Author**: [arehman626](https://github.com/arehman626)
**Repository**: [ExcelLookupC](https://github.com/arehman626/ExcelLookupC)