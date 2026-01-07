# Excel Lookup - Test Plan

## Test Environment Setup

### Prerequisites
- Windows 10/11
- .NET 8.0 Runtime installed
- Test Excel files available

### Test Data Generation
1. Run the application
2. Go to Tools > Generate Test Data
3. Select a folder for test files
4. Two files will be created:
   - `LeftFile.xlsx` (5 employee records)
   - `RightFile.xlsx` (5 staff records with some matches)

## Test Cases

### 1. Application Startup
**Objective**: Verify application starts correctly
**Steps**:
1. Run `dotnet run` or double-click `run.bat`
2. Verify main window appears
3. Verify all panels are visible and properly laid out

**Expected Result**: Application starts without errors, UI is properly rendered

### 2. File Loading - Left Panel
**Objective**: Test left file selection and loading
**Steps**:
1. Click "Select File" in left panel
2. Select `LeftFile.xlsx`
3. Verify file path appears in text box
4. Verify sheet dropdown is populated with "Employees"

**Expected Result**: File loads successfully, sheet names appear

### 3. File Loading - Right Panel
**Objective**: Test right file selection and loading
**Steps**:
1. Click "Select File" in right panel  
2. Select `RightFile.xlsx`
3. Verify file path appears in text box
4. Verify sheet dropdown is populated with "Staff"

**Expected Result**: File loads successfully, sheet names appear

### 4. Sheet Selection
**Objective**: Test sheet selection and column discovery
**Steps**:
1. Select "Employees" from left dropdown
2. Select "Staff" from right dropdown
3. Check comparison columns list

**Expected Result**: Common columns (ID, Name, Age) appear in checkbox list

### 5. Column Selection
**Objective**: Test column selection for comparison
**Steps**:
1. Check "ID" column
2. Check "Name" column
3. Verify "Compare" button becomes enabled

**Expected Result**: Selected columns are checked, Compare button enabled

### 6. Comparison Execution
**Objective**: Test the comparison process
**Steps**:
1. With ID and Name selected, click "Compare"
2. Wait for operation to complete
3. Check results display

**Expected Result**:
- Matched Records: 2 (John Smith, Alice Brown)
- Left Only Records: 3 (Jane Doe, Bob Johnson, Charlie Wilson)
- Right Only Records: 2 (David Lee, Emma Davis)

### 7. Export Functionality
**Objective**: Test result export
**Steps**:
1. After successful comparison, click "Export Results"
2. Choose save location and filename
3. Click Save
4. Verify export completes successfully

**Expected Result**: Excel file created with Summary, Matched Records, Left Only Records, and Right Only Records sheets

### 8. Export File Validation
**Objective**: Verify exported file contents
**Steps**:
1. Open exported Excel file
2. Check Summary sheet for metadata
3. Check data sheets for correct records

**Expected Result**: 
- Summary contains file names, timestamps, comparison details
- Each data sheet contains correct records with proper formatting

### 9. Error Handling - Invalid File
**Objective**: Test error handling for invalid files
**Steps**:
1. Try to select a non-Excel file (e.g., .txt file)
2. Verify appropriate error message

**Expected Result**: Clear error message, application remains stable

### 10. Error Handling - Missing Columns
**Objective**: Test error handling when no common columns exist
**Steps**:
1. Load files with no common column names
2. Attempt to compare

**Expected Result**: User is informed no common columns found

### 11. Large File Performance
**Objective**: Test with larger datasets
**Steps**:
1. Create Excel files with 1000+ rows
2. Perform comparison
3. Monitor progress bar and UI responsiveness

**Expected Result**: Application remains responsive, progress is indicated

### 12. Memory Usage
**Objective**: Test memory efficiency
**Steps**:
1. Load large files
2. Perform multiple comparisons
3. Monitor system memory usage

**Expected Result**: Memory usage remains reasonable, no memory leaks

## Regression Testing

### After Each Code Change
1. Run basic functionality tests (Cases 1-8)
2. Verify no new bugs introduced
3. Test any modified features specifically

### Before Release
1. Run complete test suite
2. Test on different Windows versions
3. Verify all error scenarios
4. Performance testing with various file sizes

## Test Data Variations

### Different File Structures
- Files with different column orders
- Files with missing columns in one sheet
- Files with extra columns
- Empty files
- Files with only headers

### Different Data Types
- Numeric IDs vs. Text IDs
- Date columns
- Boolean values
- Mixed data types

### Edge Cases
- Very large files (100k+ rows)
- Files with special characters
- Files with merged cells
- Password-protected files
- Corrupted files

## Bug Reporting Template

When reporting bugs, include:
1. **Steps to Reproduce**: Exact steps taken
2. **Expected Result**: What should happen
3. **Actual Result**: What actually happened
4. **Environment**: Windows version, .NET version
5. **Test Files**: Sample files that cause the issue
6. **Screenshots**: If applicable
7. **Error Messages**: Full error text

## Performance Benchmarks

### Acceptable Performance Targets
- File loading: < 5 seconds for files up to 10MB
- Sheet analysis: < 2 seconds for sheets up to 50,000 rows
- Comparison: < 10 seconds for comparing 10,000 rows
- Export: < 5 seconds for results up to 10,000 rows

### Memory Usage Targets
- Baseline (empty): < 50MB
- With large files loaded: < 200MB additional
- During comparison: < 300MB total
- After export: Memory released within 10 seconds