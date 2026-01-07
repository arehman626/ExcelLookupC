using ExcelLookupC.Models;
using ExcelLookupC.Services;
using ExcelLookupC.TestData;
using System.Drawing;

namespace ExcelLookupC
{
    public partial class MainForm : Form
    {
        private readonly ExcelService _excelService;
        private readonly ComparisonService _comparisonService;
        private ExcelFileInfo? _leftFileInfo;
        private ExcelFileInfo? _rightFileInfo;
        private ComparisonResult? _lastComparisonResult;
        private readonly List<ColumnMapping> _columnMappings = new();

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
            
            // Add event handlers for checkbox changes
            checkedListBoxLeftColumns.ItemCheck += CheckedListBoxColumns_ItemCheck;
            checkedListBoxRightColumns.ItemCheck += CheckedListBoxColumns_ItemCheck;
            
            // Add event handlers for new mapping controls
            buttonAddMapping.Click += ButtonAddMapping_Click;
            buttonRemoveMapping.Click += ButtonRemoveMapping_Click;
            
            // Create enhanced menu strip
            CreateMenuStrip();
            
            // Apply enhanced styling
            ApplyEnhancedStyling();
            
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
            
            checkedListBoxLeftColumns.Enabled = _leftFileInfo != null && comboBoxLeftSheet.SelectedItem != null;
            checkedListBoxRightColumns.Enabled = _rightFileInfo != null && comboBoxRightSheet.SelectedItem != null;
            
            buttonAddMapping.Enabled = checkedListBoxLeftColumns.CheckedItems.Count > 0 && 
                                     checkedListBoxRightColumns.CheckedItems.Count > 0;
            
            buttonRemoveMapping.Enabled = listBoxMappings.SelectedItem != null;
            
            buttonCompare.Enabled = _columnMappings.Any();
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

        private void ButtonAddMapping_Click(object? sender, EventArgs e)
        {
            try
            {
                var leftSelected = checkedListBoxLeftColumns.CheckedItems.Cast<string>().ToList();
                var rightSelected = checkedListBoxRightColumns.CheckedItems.Cast<string>().ToList();

                if (!leftSelected.Any() || !rightSelected.Any())
                {
                    ShowError("Please select at least one column from both left and right sheets.");
                    return;
                }

                if (leftSelected.Count != rightSelected.Count)
                {
                    ShowError("Please select the same number of columns from both sheets for mapping.");
                    return;
                }

                // Create mappings
                for (int i = 0; i < leftSelected.Count; i++)
                {
                    var mapping = new ColumnMapping
                    {
                        LeftColumn = leftSelected[i],
                        RightColumn = rightSelected[i]
                    };

                    // Check if mapping already exists
                    if (!_columnMappings.Any(m => m.LeftColumn == mapping.LeftColumn && m.RightColumn == mapping.RightColumn))
                    {
                        _columnMappings.Add(mapping);
                    }
                }

                // Clear selections
                for (int i = 0; i < checkedListBoxLeftColumns.Items.Count; i++)
                    checkedListBoxLeftColumns.SetItemChecked(i, false);
                for (int i = 0; i < checkedListBoxRightColumns.Items.Count; i++)
                    checkedListBoxRightColumns.SetItemChecked(i, false);

                UpdateMappingsDisplay();
                UpdateUI();
            }
            catch (Exception ex)
            {
                ShowError($"Error adding mapping: {ex.Message}");
            }
        }

        private void ButtonRemoveMapping_Click(object? sender, EventArgs e)
        {
            try
            {
                if (listBoxMappings.SelectedItem is ColumnMapping selectedMapping)
                {
                    _columnMappings.Remove(selectedMapping);
                    UpdateMappingsDisplay();
                    UpdateUI();
                }
                else
                {
                    ShowError("Please select a mapping to remove.");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error removing mapping: {ex.Message}");
            }
        }

        private void UpdateMappingsDisplay()
        {
            try
            {
                listBoxMappings.BeginUpdate();
                listBoxMappings.DataSource = null;
                
                if (_columnMappings.Any())
                {
                    listBoxMappings.DataSource = new List<ColumnMapping>(_columnMappings);
                    listBoxMappings.DisplayMember = nameof(ColumnMapping.DisplayText);
                }
                else
                {
                    listBoxMappings.Items.Clear();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error updating mappings display: {ex.Message}");
            }
            finally
            {
                listBoxMappings.EndUpdate();
            }
        }

        private async Task UpdateAvailableColumnsAsync()
        {
            try
            {
                await Task.Run(() =>
                {
                    this.Invoke(() =>
                    {
                        // Clear existing items and mappings
                        checkedListBoxLeftColumns.Items.Clear();
                        checkedListBoxRightColumns.Items.Clear();
                        _columnMappings.Clear();
                        UpdateMappingsDisplay();
                        
                        // Populate left columns
                        if (_leftFileInfo != null && comboBoxLeftSheet.SelectedItem != null)
                        {
                            var leftSheetName = comboBoxLeftSheet.SelectedItem.ToString()!;
                            var leftColumns = _leftFileInfo.SheetColumns[leftSheetName];
                            foreach (var column in leftColumns.OrderBy(c => c))
                            {
                                checkedListBoxLeftColumns.Items.Add(column);
                            }
                        }
                        
                        // Populate right columns
                        if (_rightFileInfo != null && comboBoxRightSheet.SelectedItem != null)
                        {
                            var rightSheetName = comboBoxRightSheet.SelectedItem.ToString()!;
                            var rightColumns = _rightFileInfo.SheetColumns[rightSheetName];
                            foreach (var column in rightColumns.OrderBy(c => c))
                            {
                                checkedListBoxRightColumns.Items.Add(column);
                            }
                        }
                        
                        var leftCount = checkedListBoxLeftColumns.Items.Count;
                        var rightCount = checkedListBoxRightColumns.Items.Count;
                        UpdateStatus($"Loaded {leftCount} left columns and {rightCount} right columns. Select columns and create mappings.");
                        
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
                !_columnMappings.Any())
            {
                ShowError("Please ensure both files are loaded, sheets are selected, and column mappings are created.");
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

                UpdateStatus("Comparing data...");

                _lastComparisonResult = _comparisonService.CompareSheetsWithMapping(
                    leftSheetData, rightSheetData, _columnMappings,
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
                    
                    var result = ShowInfoWithOptions($"Comparison results have been successfully exported to:\n{saveFileDialog1.FileName}\n\nWould you like to open the file now?",
                        "Export Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    
                    if (result == DialogResult.Yes)
                    {
                        OpenExportedFile(saveFileDialog1.FileName);
                    }
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

        private DialogResult ShowInfoWithOptions(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(message, title, buttons, icon);
        }

        private void OpenExportedFile(string filePath)
        {
            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(startInfo);
                UpdateStatus($"Opened exported file: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                ShowError($"Error opening file: {ex.Message}\n\nFile location: {filePath}");
            }
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

        private void CreateMenuStrip()
        {
            var menuStrip = new MenuStrip();
            
            // Enhance menu strip styling
            menuStrip.BackColor = Color.FromArgb(70, 130, 180);
            menuStrip.ForeColor = Color.White;
            menuStrip.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            
            // Tools menu
            var toolsMenu = new ToolStripMenuItem("Tools");
            toolsMenu.ForeColor = Color.White;
            
            var generateTestDataItem = new ToolStripMenuItem("Generate Test Data");
            generateTestDataItem.Click += async (s, e) => await GenerateTestDataAsync();
            toolsMenu.DropDownItems.Add(generateTestDataItem);
            
            menuStrip.Items.Add(toolsMenu);
            
            // Help menu
            var helpMenu = new ToolStripMenuItem("Help");
            helpMenu.ForeColor = Color.White;
            
            var aboutItem = new ToolStripMenuItem("About Excel Lookup...");
            aboutItem.Click += (s, e) => ShowAboutDialog();
            helpMenu.DropDownItems.Add(aboutItem);
            
            menuStrip.Items.Add(helpMenu);
            
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        private void ApplyEnhancedStyling()
        {
            // Form styling
            this.BackColor = Color.FromArgb(240, 248, 255);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            
            // Group box styling
            var groupBoxes = new[] { groupBoxLeft, groupBoxRight, groupBoxComparison, groupBoxResults };
            foreach (var groupBox in groupBoxes)
            {
                groupBox.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                groupBox.ForeColor = Color.FromArgb(70, 130, 180);
                groupBox.BackColor = Color.White;
            }
            
            // Button styling
            var primaryButtons = new[] { buttonSelectLeftFile, buttonSelectRightFile, buttonCompare, buttonExport };
            foreach (var button in primaryButtons)
            {
                button.BackColor = Color.FromArgb(70, 130, 180);
                button.ForeColor = Color.White;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            }
            
            // Secondary button styling
            var secondaryButtons = new[] { buttonAddMapping, buttonRemoveMapping };
            foreach (var button in secondaryButtons)
            {
                button.BackColor = Color.FromArgb(100, 160, 200);
                button.ForeColor = Color.White;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
            }
            
            // Text box styling
            var textBoxes = new[] { textBoxLeftFile, textBoxRightFile, textBoxMatched, textBoxLeftOnly, textBoxRightOnly };
            foreach (var textBox in textBoxes)
            {
                textBox.BackColor = Color.FromArgb(250, 250, 250);
                textBox.BorderStyle = BorderStyle.FixedSingle;
                textBox.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            }
            
            // ComboBox styling
            var comboBoxes = new[] { comboBoxLeftSheet, comboBoxRightSheet };
            foreach (var comboBox in comboBoxes)
            {
                comboBox.BackColor = Color.White;
                comboBox.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            }
            
            // CheckedListBox styling
            var checkedListBoxes = new[] { checkedListBoxLeftColumns, checkedListBoxRightColumns };
            foreach (var clb in checkedListBoxes)
            {
                clb.BackColor = Color.White;
                clb.BorderStyle = BorderStyle.FixedSingle;
                clb.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
            }
            
            // ListBox styling
            listBoxMappings.BackColor = Color.White;
            listBoxMappings.BorderStyle = BorderStyle.FixedSingle;
            listBoxMappings.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
            
            // ProgressBar styling
            progressBar1.BackColor = Color.FromArgb(230, 240, 250);
            
            // Status strip styling
            statusStrip1.BackColor = Color.FromArgb(70, 130, 180);
            statusStrip1.ForeColor = Color.White;
            toolStripStatusLabel1.ForeColor = Color.White;
        }

        private void ShowAboutDialog()
        {
            using var aboutForm = new AboutForm();
            aboutForm.ShowDialog(this);
        }
    }
}