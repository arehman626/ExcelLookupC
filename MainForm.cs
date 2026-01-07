using ExcelLookupC.Models;
using ExcelLookupC.Services;
using ExcelLookupC.TestData;

namespace ExcelLookupC
{
    public partial class MainForm : Form
    {
        private readonly ExcelService _excelService;
        private readonly ComparisonService _comparisonService;
        private ExcelFileInfo? _leftFileInfo;
        private ExcelFileInfo? _rightFileInfo;
        private ComparisonResult? _lastComparisonResult;

        public MainForm()
        {
            InitializeComponent();
            _excelService = new ExcelService();
            _comparisonService = new ComparisonService();
            InitializeUI();
        }



        private void InitializeUI()
        {
            UpdateStatus("Ready to load Excel files.");
            
            // Add event handler for checkbox changes
            checkedListBoxColumns.ItemCheck += CheckedListBoxColumns_ItemCheck;
            
            // Add a menu strip for additional options
            var menuStrip = new MenuStrip();
            var toolsMenu = new ToolStripMenuItem("Tools");
            var generateTestDataItem = new ToolStripMenuItem("Generate Test Data");
            generateTestDataItem.Click += async (s, e) => await GenerateTestDataAsync();
            toolsMenu.DropDownItems.Add(generateTestDataItem);
            menuStrip.Items.Add(toolsMenu);
            
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
            
            UpdateUI();
        }

        private void UpdateStatus(string message)
        {
            toolStripStatusLabel1.Text = message;
            statusStrip1.Refresh();
        }

        private void UpdateUI()
        {
            // Enable/disable controls based on current state
            comboBoxLeftSheet.Enabled = _leftFileInfo != null;
            comboBoxRightSheet.Enabled = _rightFileInfo != null;
            
            bool canCompare = _leftFileInfo != null && _rightFileInfo != null &&
                            comboBoxLeftSheet.SelectedItem != null && comboBoxRightSheet.SelectedItem != null &&
                            checkedListBoxColumns.CheckedItems.Count > 0;
            
            buttonCompare.Enabled = canCompare;
            buttonExport.Enabled = _lastComparisonResult != null;
        }

        private async void ButtonSelectLeftFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    await LoadLeftFileAsync(openFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading left file: {ex.Message}");
            }
        }

        private async void ButtonSelectRightFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    await LoadRightFileAsync(openFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading right file: {ex.Message}");
            }
        }

        private async Task LoadLeftFileAsync(string filePath)
        {
            progressBar1.Style = ProgressBarStyle.Marquee;
            UpdateStatus("Loading left file...");

            try
            {
                _leftFileInfo = await _excelService.LoadExcelFileAsync(filePath);
                textBoxLeftFile.Text = _leftFileInfo.FileName;
                
                comboBoxLeftSheet.DataSource = _leftFileInfo.SheetNames;
                comboBoxLeftSheet.SelectedIndex = _leftFileInfo.SheetNames.Count > 0 ? 0 : -1;

                UpdateStatus($"Loaded left file: {_leftFileInfo.FileName} ({_leftFileInfo.SheetNames.Count} sheets)");
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                UpdateUI();
            }
        }

        private async Task LoadRightFileAsync(string filePath)
        {
            progressBar1.Style = ProgressBarStyle.Marquee;
            UpdateStatus("Loading right file...");

            try
            {
                _rightFileInfo = await _excelService.LoadExcelFileAsync(filePath);
                textBoxRightFile.Text = _rightFileInfo.FileName;
                
                comboBoxRightSheet.DataSource = _rightFileInfo.SheetNames;
                comboBoxRightSheet.SelectedIndex = _rightFileInfo.SheetNames.Count > 0 ? 0 : -1;

                UpdateStatus($"Loaded right file: {_rightFileInfo.FileName} ({_rightFileInfo.SheetNames.Count} sheets)");
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                UpdateUI();
            }
        }

        private async void ComboBoxLeftSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            await UpdateAvailableColumnsAsync();
        }

        private async void ComboBoxRightSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            await UpdateAvailableColumnsAsync();
        }

        private void CheckedListBoxColumns_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            // Use BeginInvoke to update UI after the check state has changed
            this.BeginInvoke(new Action(UpdateUI));
        }

        private async Task UpdateAvailableColumnsAsync()
        {
            if (_leftFileInfo == null || _rightFileInfo == null ||
                comboBoxLeftSheet.SelectedItem == null || comboBoxRightSheet.SelectedItem == null)
            {
                checkedListBoxColumns.Items.Clear();
                UpdateUI();
                return;
            }

            try
            {
                await Task.Run(() =>
                {
                    var leftSheetName = comboBoxLeftSheet.SelectedItem.ToString()!;
                    var rightSheetName = comboBoxRightSheet.SelectedItem.ToString()!;

                    var leftColumns = _leftFileInfo.SheetColumns[leftSheetName];
                    var rightColumns = _rightFileInfo.SheetColumns[rightSheetName];

                    // Find common columns
                    var commonColumns = leftColumns.Intersect(rightColumns, StringComparer.OrdinalIgnoreCase).ToList();

                    this.Invoke(() =>
                    {
                        checkedListBoxColumns.Items.Clear();
                        foreach (var column in commonColumns.OrderBy(c => c))
                        {
                            checkedListBoxColumns.Items.Add(column);
                        }

                        if (commonColumns.Any())
                        {
                            UpdateStatus($"Found {commonColumns.Count} common columns between sheets.");
                        }
                        else
                        {
                            UpdateStatus("No common columns found between selected sheets.");
                        }
                        
                        // Update UI state after populating columns
                        UpdateUI();
                    });
                });
            }
            catch (Exception ex)
            {
                ShowError($"Error updating columns: {ex.Message}");
            }
            finally
            {
                UpdateUI();
            }
        }

        private async void ButtonCompare_Click(object sender, EventArgs e)
        {
            if (_leftFileInfo == null || _rightFileInfo == null ||
                comboBoxLeftSheet.SelectedItem == null || comboBoxRightSheet.SelectedItem == null ||
                checkedListBoxColumns.CheckedItems.Count == 0)
            {
                ShowError("Please ensure both files are loaded, sheets are selected, and comparison columns are checked.");
                return;
            }

            progressBar1.Style = ProgressBarStyle.Marquee;
            buttonCompare.Enabled = false;

            try
            {
                UpdateStatus("Loading sheet data...");

                var leftSheetName = comboBoxLeftSheet.SelectedItem.ToString()!;
                var rightSheetName = comboBoxRightSheet.SelectedItem.ToString()!;

                var leftSheetData = await _excelService.LoadSheetDataAsync(_leftFileInfo.FilePath, leftSheetName);
                var rightSheetData = await _excelService.LoadSheetDataAsync(_rightFileInfo.FilePath, rightSheetName);

                var comparisonColumns = checkedListBoxColumns.CheckedItems.Cast<string>().ToList();

                UpdateStatus("Comparing data...");

                _lastComparisonResult = _comparisonService.CompareSheets(
                    leftSheetData, rightSheetData, comparisonColumns,
                    _leftFileInfo.FileName, _rightFileInfo.FileName);

                // Update results display
                textBoxMatched.Text = _lastComparisonResult.MatchedRecords.Count.ToString();
                textBoxLeftOnly.Text = _lastComparisonResult.LeftOnlyRecords.Count.ToString();
                textBoxRightOnly.Text = _lastComparisonResult.RightOnlyRecords.Count.ToString();

                UpdateStatus($"Comparison complete. Found {_lastComparisonResult.MatchedRecords.Count} matches, " +
                           $"{_lastComparisonResult.LeftOnlyRecords.Count} left-only, " +
                           $"{_lastComparisonResult.RightOnlyRecords.Count} right-only records.");

                ShowInfo($"Comparison completed successfully!\n\n" +
                        $"Results:\n" +
                        $"• Matched Records: {_lastComparisonResult.MatchedRecords.Count}\n" +
                        $"• Left Only Records: {_lastComparisonResult.LeftOnlyRecords.Count}\n" +
                        $"• Right Only Records: {_lastComparisonResult.RightOnlyRecords.Count}");
            }
            catch (Exception ex)
            {
                ShowError($"Error during comparison: {ex.Message}");
                _lastComparisonResult = null;
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                buttonCompare.Enabled = true;
                UpdateUI();
            }
        }

        private async void ButtonExport_Click(object sender, EventArgs e)
        {
            if (_lastComparisonResult == null)
            {
                ShowError("No comparison results to export.");
                return;
            }

            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var defaultFileName = $"ExcelComparison_{timestamp}.xlsx";
                
                saveFileDialog1.FileName = defaultFileName;
                
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    buttonExport.Enabled = false;
                    UpdateStatus("Exporting results...");

                    await _excelService.ExportComparisonResultAsync(_lastComparisonResult, saveFileDialog1.FileName);

                    UpdateStatus($"Results exported to: {saveFileDialog1.FileName}");
                    ShowInfo($"Comparison results have been successfully exported to:\n{saveFileDialog1.FileName}");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error exporting results: {ex.Message}");
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                buttonExport.Enabled = true;
                UpdateUI();
            }
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowInfo(string message)
        {
            MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private async Task GenerateTestDataAsync()
        {
            try
            {
                var folderBrowser = new FolderBrowserDialog();
                folderBrowser.Description = "Select folder to create test Excel files";
                folderBrowser.ShowNewFolderButton = true;

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    UpdateStatus("Generating test data...");

                    await TestDataGenerator.CreateSampleFilesAsync(folderBrowser.SelectedPath);

                    UpdateStatus($"Test files created in: {folderBrowser.SelectedPath}");
                    ShowInfo($"Test files have been created:\n• LeftFile.xlsx\n• RightFile.xlsx\n\nLocation: {folderBrowser.SelectedPath}");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error generating test data: {ex.Message}");
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
            }
        }
    }
}