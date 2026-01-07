using System.Drawing;

namespace ExcelLookupC
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new TableLayoutPanel();
            this.groupBoxLeft = new GroupBox();
            this.tableLayoutPanelLeft = new TableLayoutPanel();
            this.labelLeftFile = new Label();
            this.buttonSelectLeftFile = new Button();
            this.textBoxLeftFile = new TextBox();
            this.labelLeftSheet = new Label();
            this.comboBoxLeftSheet = new ComboBox();
            this.groupBoxRight = new GroupBox();
            this.tableLayoutPanelRight = new TableLayoutPanel();
            this.labelRightFile = new Label();
            this.buttonSelectRightFile = new Button();
            this.textBoxRightFile = new TextBox();
            this.labelRightSheet = new Label();
            this.comboBoxRightSheet = new ComboBox();
            this.groupBoxComparison = new GroupBox();
            this.tableLayoutPanelComparison = new TableLayoutPanel();
            this.labelLeftColumns = new Label();
            this.checkedListBoxLeftColumns = new CheckedListBox();
            this.labelRightColumns = new Label();
            this.checkedListBoxRightColumns = new CheckedListBox();
            this.buttonAddMapping = new Button();
            this.labelMappings = new Label();
            this.listBoxMappings = new ListBox();
            this.buttonRemoveMapping = new Button();
            this.buttonCompare = new Button();
            this.buttonExport = new Button();
            this.groupBoxResults = new GroupBox();
            this.tableLayoutPanelResults = new TableLayoutPanel();
            this.labelMatched = new Label();
            this.labelLeftOnly = new Label();
            this.labelRightOnly = new Label();
            this.textBoxMatched = new TextBox();
            this.textBoxLeftOnly = new TextBox();
            this.textBoxRightOnly = new TextBox();
            this.progressBar1 = new ProgressBar();
            this.statusStrip1 = new StatusStrip();
            this.toolStripStatusLabel1 = new ToolStripStatusLabel();
            this.openFileDialog1 = new OpenFileDialog();
            this.saveFileDialog1 = new SaveFileDialog();
            
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBoxLeft.SuspendLayout();
            this.tableLayoutPanelLeft.SuspendLayout();
            this.groupBoxRight.SuspendLayout();
            this.tableLayoutPanelRight.SuspendLayout();
            this.groupBoxComparison.SuspendLayout();
            this.tableLayoutPanelComparison.SuspendLayout();
            this.groupBoxResults.SuspendLayout();
            this.tableLayoutPanelResults.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.groupBoxLeft, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.groupBoxRight, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.groupBoxComparison, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.groupBoxResults, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.progressBar1, 0, 2);
            this.tableLayoutPanel1.Dock = DockStyle.Fill;
            this.tableLayoutPanel1.Location = new Point(0, 28);
            this.tableLayoutPanel1.Margin = new Padding(5);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new Padding(10);
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 10F));
            this.tableLayoutPanel1.Size = new Size(1200, 746);
            this.tableLayoutPanel1.TabIndex = 0;
            
            // 
            // groupBoxLeft
            // 
            this.groupBoxLeft.Controls.Add(this.tableLayoutPanelLeft);
            this.groupBoxLeft.Dock = DockStyle.Fill;
            this.groupBoxLeft.Location = new Point(3, 3);
            this.groupBoxLeft.Name = "groupBoxLeft";
            this.groupBoxLeft.Size = new Size(494, 309);
            this.groupBoxLeft.TabIndex = 0;
            this.groupBoxLeft.TabStop = false;
            this.groupBoxLeft.Text = "Left File";
            
            // 
            // tableLayoutPanelLeft
            // 
            this.tableLayoutPanelLeft.ColumnCount = 2;
            this.tableLayoutPanelLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            this.tableLayoutPanelLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80F));
            this.tableLayoutPanelLeft.Controls.Add(this.labelLeftFile, 0, 0);
            this.tableLayoutPanelLeft.Controls.Add(this.textBoxLeftFile, 1, 0);
            this.tableLayoutPanelLeft.Controls.Add(this.buttonSelectLeftFile, 0, 1);
            this.tableLayoutPanelLeft.Controls.Add(this.labelLeftSheet, 0, 2);
            this.tableLayoutPanelLeft.Controls.Add(this.comboBoxLeftSheet, 1, 2);
            this.tableLayoutPanelLeft.Dock = DockStyle.Fill;
            this.tableLayoutPanelLeft.Location = new Point(3, 23);
            this.tableLayoutPanelLeft.Name = "tableLayoutPanelLeft";
            this.tableLayoutPanelLeft.RowCount = 4;
            this.tableLayoutPanelLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            this.tableLayoutPanelLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.tableLayoutPanelLeft.Size = new Size(488, 283);
            this.tableLayoutPanelLeft.TabIndex = 0;
            
            // 
            // labelLeftFile
            // 
            this.labelLeftFile.AutoSize = true;
            this.labelLeftFile.Dock = DockStyle.Fill;
            this.labelLeftFile.Location = new Point(3, 0);
            this.labelLeftFile.Name = "labelLeftFile";
            this.labelLeftFile.Size = new Size(91, 30);
            this.labelLeftFile.TabIndex = 0;
            this.labelLeftFile.Text = "File:";
            this.labelLeftFile.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // textBoxLeftFile
            // 
            this.textBoxLeftFile.Dock = DockStyle.Fill;
            this.textBoxLeftFile.Location = new Point(100, 3);
            this.textBoxLeftFile.Name = "textBoxLeftFile";
            this.textBoxLeftFile.ReadOnly = true;
            this.textBoxLeftFile.Size = new Size(385, 27);
            this.textBoxLeftFile.TabIndex = 1;
            
            // 
            // buttonSelectLeftFile
            // 
            this.buttonSelectLeftFile.Dock = DockStyle.Fill;
            this.buttonSelectLeftFile.Location = new Point(3, 33);
            this.buttonSelectLeftFile.Name = "buttonSelectLeftFile";
            this.buttonSelectLeftFile.Size = new Size(91, 29);
            this.buttonSelectLeftFile.TabIndex = 2;
            this.buttonSelectLeftFile.Text = "Select File";
            this.buttonSelectLeftFile.UseVisualStyleBackColor = true;
            this.buttonSelectLeftFile.Click += new EventHandler(this.ButtonSelectLeftFile_Click);
            
            // 
            // labelLeftSheet
            // 
            this.labelLeftSheet.AutoSize = true;
            this.labelLeftSheet.Dock = DockStyle.Fill;
            this.labelLeftSheet.Location = new Point(3, 65);
            this.labelLeftSheet.Name = "labelLeftSheet";
            this.labelLeftSheet.Size = new Size(91, 30);
            this.labelLeftSheet.TabIndex = 3;
            this.labelLeftSheet.Text = "Sheet:";
            this.labelLeftSheet.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // comboBoxLeftSheet
            // 
            this.comboBoxLeftSheet.Dock = DockStyle.Fill;
            this.comboBoxLeftSheet.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBoxLeftSheet.Location = new Point(100, 68);
            this.comboBoxLeftSheet.Name = "comboBoxLeftSheet";
            this.comboBoxLeftSheet.Size = new Size(385, 28);
            this.comboBoxLeftSheet.TabIndex = 4;
            this.comboBoxLeftSheet.SelectedIndexChanged += new EventHandler(this.ComboBoxLeftSheet_SelectedIndexChanged);
            
            // 
            // groupBoxRight
            // 
            this.groupBoxRight.Controls.Add(this.tableLayoutPanelRight);
            this.groupBoxRight.Dock = DockStyle.Fill;
            this.groupBoxRight.Location = new Point(503, 3);
            this.groupBoxRight.Name = "groupBoxRight";
            this.groupBoxRight.Size = new Size(494, 309);
            this.groupBoxRight.TabIndex = 1;
            this.groupBoxRight.TabStop = false;
            this.groupBoxRight.Text = "Right File";
            
            // 
            // tableLayoutPanelRight
            // 
            this.tableLayoutPanelRight.ColumnCount = 2;
            this.tableLayoutPanelRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            this.tableLayoutPanelRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80F));
            this.tableLayoutPanelRight.Controls.Add(this.labelRightFile, 0, 0);
            this.tableLayoutPanelRight.Controls.Add(this.textBoxRightFile, 1, 0);
            this.tableLayoutPanelRight.Controls.Add(this.buttonSelectRightFile, 0, 1);
            this.tableLayoutPanelRight.Controls.Add(this.labelRightSheet, 0, 2);
            this.tableLayoutPanelRight.Controls.Add(this.comboBoxRightSheet, 1, 2);
            this.tableLayoutPanelRight.Dock = DockStyle.Fill;
            this.tableLayoutPanelRight.Location = new Point(3, 23);
            this.tableLayoutPanelRight.Name = "tableLayoutPanelRight";
            this.tableLayoutPanelRight.RowCount = 4;
            this.tableLayoutPanelRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            this.tableLayoutPanelRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelRight.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.tableLayoutPanelRight.Size = new Size(488, 283);
            this.tableLayoutPanelRight.TabIndex = 0;
            
            // 
            // labelRightFile
            // 
            this.labelRightFile.AutoSize = true;
            this.labelRightFile.Dock = DockStyle.Fill;
            this.labelRightFile.Location = new Point(3, 0);
            this.labelRightFile.Name = "labelRightFile";
            this.labelRightFile.Size = new Size(91, 30);
            this.labelRightFile.TabIndex = 0;
            this.labelRightFile.Text = "File:";
            this.labelRightFile.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // textBoxRightFile
            // 
            this.textBoxRightFile.Dock = DockStyle.Fill;
            this.textBoxRightFile.Location = new Point(100, 3);
            this.textBoxRightFile.Name = "textBoxRightFile";
            this.textBoxRightFile.ReadOnly = true;
            this.textBoxRightFile.Size = new Size(385, 27);
            this.textBoxRightFile.TabIndex = 1;
            
            // 
            // buttonSelectRightFile
            // 
            this.buttonSelectRightFile.Dock = DockStyle.Fill;
            this.buttonSelectRightFile.Location = new Point(3, 33);
            this.buttonSelectRightFile.Name = "buttonSelectRightFile";
            this.buttonSelectRightFile.Size = new Size(91, 29);
            this.buttonSelectRightFile.TabIndex = 2;
            this.buttonSelectRightFile.Text = "Select File";
            this.buttonSelectRightFile.UseVisualStyleBackColor = true;
            this.buttonSelectRightFile.Click += new EventHandler(this.ButtonSelectRightFile_Click);
            
            // 
            // labelRightSheet
            // 
            this.labelRightSheet.AutoSize = true;
            this.labelRightSheet.Dock = DockStyle.Fill;
            this.labelRightSheet.Location = new Point(3, 65);
            this.labelRightSheet.Name = "labelRightSheet";
            this.labelRightSheet.Size = new Size(91, 30);
            this.labelRightSheet.TabIndex = 3;
            this.labelRightSheet.Text = "Sheet:";
            this.labelRightSheet.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // comboBoxRightSheet
            // 
            this.comboBoxRightSheet.Dock = DockStyle.Fill;
            this.comboBoxRightSheet.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBoxRightSheet.Location = new Point(100, 68);
            this.comboBoxRightSheet.Name = "comboBoxRightSheet";
            this.comboBoxRightSheet.Size = new Size(385, 28);
            this.comboBoxRightSheet.TabIndex = 4;
            this.comboBoxRightSheet.SelectedIndexChanged += new EventHandler(this.ComboBoxRightSheet_SelectedIndexChanged);
            
            // 
            // groupBoxComparison
            // 
            this.groupBoxComparison.Controls.Add(this.tableLayoutPanelComparison);
            this.groupBoxComparison.Dock = DockStyle.Fill;
            this.groupBoxComparison.Location = new Point(3, 318);
            this.groupBoxComparison.Name = "groupBoxComparison";
            this.groupBoxComparison.Size = new Size(494, 309);
            this.groupBoxComparison.TabIndex = 2;
            this.groupBoxComparison.TabStop = false;
            this.groupBoxComparison.Text = "Column Mapping & Comparison";
            
            // 
            // tableLayoutPanelComparison
            // 
            this.tableLayoutPanelComparison.ColumnCount = 3;
            this.tableLayoutPanelComparison.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            this.tableLayoutPanelComparison.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            this.tableLayoutPanelComparison.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            this.tableLayoutPanelComparison.Controls.Add(this.labelLeftColumns, 0, 0);
            this.tableLayoutPanelComparison.Controls.Add(this.labelRightColumns, 2, 0);
            this.tableLayoutPanelComparison.Controls.Add(this.checkedListBoxLeftColumns, 0, 1);
            this.tableLayoutPanelComparison.Controls.Add(this.checkedListBoxRightColumns, 2, 1);
            this.tableLayoutPanelComparison.Controls.Add(this.buttonAddMapping, 1, 1);
            this.tableLayoutPanelComparison.Controls.Add(this.labelMappings, 0, 2);
            this.tableLayoutPanelComparison.Controls.Add(this.listBoxMappings, 0, 3);
            this.tableLayoutPanelComparison.Controls.Add(this.buttonRemoveMapping, 1, 3);
            this.tableLayoutPanelComparison.Controls.Add(this.buttonCompare, 0, 4);
            this.tableLayoutPanelComparison.Dock = DockStyle.Fill;
            this.tableLayoutPanelComparison.Location = new Point(3, 23);
            this.tableLayoutPanelComparison.Name = "tableLayoutPanelComparison";
            this.tableLayoutPanelComparison.RowCount = 5;
            this.tableLayoutPanelComparison.RowStyles.Add(new RowStyle(SizeType.Absolute, 25F));
            this.tableLayoutPanelComparison.RowStyles.Add(new RowStyle(SizeType.Percent, 40F));
            this.tableLayoutPanelComparison.RowStyles.Add(new RowStyle(SizeType.Absolute, 25F));
            this.tableLayoutPanelComparison.RowStyles.Add(new RowStyle(SizeType.Percent, 40F));
            this.tableLayoutPanelComparison.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            this.tableLayoutPanelComparison.Size = new Size(488, 283);
            this.tableLayoutPanelComparison.TabIndex = 0;
            
            // 
            // labelLeftColumns
            // 
            this.labelLeftColumns.AutoSize = true;
            this.labelLeftColumns.Dock = DockStyle.Fill;
            this.labelLeftColumns.Location = new Point(3, 0);
            this.labelLeftColumns.Name = "labelLeftColumns";
            this.labelLeftColumns.Size = new Size(189, 25);
            this.labelLeftColumns.TabIndex = 0;
            this.labelLeftColumns.Text = "Left Sheet Columns:";
            this.labelLeftColumns.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // labelRightColumns
            // 
            this.labelRightColumns.AutoSize = true;
            this.labelRightColumns.Dock = DockStyle.Fill;
            this.labelRightColumns.Location = new Point(295, 0);
            this.labelRightColumns.Name = "labelRightColumns";
            this.labelRightColumns.Size = new Size(190, 25);
            this.labelRightColumns.TabIndex = 1;
            this.labelRightColumns.Text = "Right Sheet Columns:";
            this.labelRightColumns.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // checkedListBoxLeftColumns
            // 
            this.checkedListBoxLeftColumns.Dock = DockStyle.Fill;
            this.checkedListBoxLeftColumns.FormattingEnabled = true;
            this.checkedListBoxLeftColumns.Location = new Point(3, 28);
            this.checkedListBoxLeftColumns.Name = "checkedListBoxLeftColumns";
            this.checkedListBoxLeftColumns.Size = new Size(189, 87);
            this.checkedListBoxLeftColumns.TabIndex = 2;
            
            // 
            // checkedListBoxRightColumns
            // 
            this.checkedListBoxRightColumns.Dock = DockStyle.Fill;
            this.checkedListBoxRightColumns.FormattingEnabled = true;
            this.checkedListBoxRightColumns.Location = new Point(295, 28);
            this.checkedListBoxRightColumns.Name = "checkedListBoxRightColumns";
            this.checkedListBoxRightColumns.Size = new Size(190, 87);
            this.checkedListBoxRightColumns.TabIndex = 3;
            
            // 
            // buttonAddMapping
            // 
            this.buttonAddMapping.Dock = DockStyle.Fill;
            this.buttonAddMapping.Location = new Point(198, 35);
            this.buttonAddMapping.Name = "buttonAddMapping";
            this.buttonAddMapping.Size = new Size(91, 73);
            this.buttonAddMapping.TabIndex = 4;
            this.buttonAddMapping.Text = "Add\nMapping\n→";
            this.buttonAddMapping.UseVisualStyleBackColor = true;
            this.buttonAddMapping.Click += new EventHandler(this.ButtonAddMapping_Click);
            
            // 
            // labelMappings
            // 
            this.labelMappings.AutoSize = true;
            this.labelMappings.Dock = DockStyle.Fill;
            this.labelMappings.Location = new Point(3, 118);
            this.labelMappings.Name = "labelMappings";
            this.labelMappings.Size = new Size(189, 25);
            this.labelMappings.TabIndex = 5;
            this.labelMappings.Text = "Column Mappings:";
            this.labelMappings.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // listBoxMappings
            // 
            this.tableLayoutPanelComparison.SetColumnSpan(this.listBoxMappings, 2);
            this.listBoxMappings.Dock = DockStyle.Fill;
            this.listBoxMappings.FormattingEnabled = true;
            this.listBoxMappings.Location = new Point(3, 146);
            this.listBoxMappings.Name = "listBoxMappings";
            this.listBoxMappings.Size = new Size(286, 87);
            this.listBoxMappings.TabIndex = 6;
            
            // 
            // buttonRemoveMapping
            // 
            this.buttonRemoveMapping.Dock = DockStyle.Fill;
            this.buttonRemoveMapping.Location = new Point(295, 146);
            this.buttonRemoveMapping.Name = "buttonRemoveMapping";
            this.buttonRemoveMapping.Size = new Size(190, 87);
            this.buttonRemoveMapping.TabIndex = 7;
            this.buttonRemoveMapping.Text = "Remove\nSelected\nMapping";
            this.buttonRemoveMapping.UseVisualStyleBackColor = true;
            this.buttonRemoveMapping.Click += new EventHandler(this.ButtonRemoveMapping_Click);
            
            // 
            // buttonCompare
            // 
            this.tableLayoutPanelComparison.SetColumnSpan(this.buttonCompare, 3);
            this.buttonCompare.Dock = DockStyle.Fill;
            this.buttonCompare.Enabled = false;
            this.buttonCompare.Location = new Point(3, 239);
            this.buttonCompare.Name = "buttonCompare";
            this.buttonCompare.Size = new Size(482, 34);
            this.buttonCompare.TabIndex = 8;
            this.buttonCompare.Text = "Compare Mapped Columns";
            this.buttonCompare.UseVisualStyleBackColor = true;
            this.buttonCompare.Click += new EventHandler(this.ButtonCompare_Click);
            
            // 
            // groupBoxResults
            // 
            this.groupBoxResults.Controls.Add(this.tableLayoutPanelResults);
            this.groupBoxResults.Dock = DockStyle.Fill;
            this.groupBoxResults.Location = new Point(503, 318);
            this.groupBoxResults.Name = "groupBoxResults";
            this.groupBoxResults.Size = new Size(494, 309);
            this.groupBoxResults.TabIndex = 3;
            this.groupBoxResults.TabStop = false;
            this.groupBoxResults.Text = "Comparison Results";
            
            // 
            // tableLayoutPanelResults
            // 
            this.tableLayoutPanelResults.ColumnCount = 2;
            this.tableLayoutPanelResults.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            this.tableLayoutPanelResults.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            this.tableLayoutPanelResults.Controls.Add(this.labelMatched, 0, 0);
            this.tableLayoutPanelResults.Controls.Add(this.textBoxMatched, 1, 0);
            this.tableLayoutPanelResults.Controls.Add(this.labelLeftOnly, 0, 1);
            this.tableLayoutPanelResults.Controls.Add(this.textBoxLeftOnly, 1, 1);
            this.tableLayoutPanelResults.Controls.Add(this.labelRightOnly, 0, 2);
            this.tableLayoutPanelResults.Controls.Add(this.textBoxRightOnly, 1, 2);
            this.tableLayoutPanelResults.Controls.Add(this.buttonExport, 0, 3);
            this.tableLayoutPanelResults.Dock = DockStyle.Fill;
            this.tableLayoutPanelResults.Location = new Point(3, 23);
            this.tableLayoutPanelResults.Name = "tableLayoutPanelResults";
            this.tableLayoutPanelResults.RowCount = 5;
            this.tableLayoutPanelResults.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelResults.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelResults.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            this.tableLayoutPanelResults.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            this.tableLayoutPanelResults.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.tableLayoutPanelResults.Size = new Size(488, 283);
            this.tableLayoutPanelResults.TabIndex = 0;
            
            // 
            // labelMatched
            // 
            this.labelMatched.AutoSize = true;
            this.labelMatched.Dock = DockStyle.Fill;
            this.labelMatched.Location = new Point(3, 0);
            this.labelMatched.Name = "labelMatched";
            this.labelMatched.Size = new Size(140, 30);
            this.labelMatched.TabIndex = 0;
            this.labelMatched.Text = "Matched Records:";
            this.labelMatched.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // textBoxMatched
            // 
            this.textBoxMatched.Dock = DockStyle.Fill;
            this.textBoxMatched.Location = new Point(149, 3);
            this.textBoxMatched.Name = "textBoxMatched";
            this.textBoxMatched.ReadOnly = true;
            this.textBoxMatched.Size = new Size(336, 27);
            this.textBoxMatched.TabIndex = 1;
            
            // 
            // labelLeftOnly
            // 
            this.labelLeftOnly.AutoSize = true;
            this.labelLeftOnly.Dock = DockStyle.Fill;
            this.labelLeftOnly.Location = new Point(3, 30);
            this.labelLeftOnly.Name = "labelLeftOnly";
            this.labelLeftOnly.Size = new Size(140, 30);
            this.labelLeftOnly.TabIndex = 2;
            this.labelLeftOnly.Text = "Left Only Records:";
            this.labelLeftOnly.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // textBoxLeftOnly
            // 
            this.textBoxLeftOnly.Dock = DockStyle.Fill;
            this.textBoxLeftOnly.Location = new Point(149, 33);
            this.textBoxLeftOnly.Name = "textBoxLeftOnly";
            this.textBoxLeftOnly.ReadOnly = true;
            this.textBoxLeftOnly.Size = new Size(336, 27);
            this.textBoxLeftOnly.TabIndex = 3;
            
            // 
            // labelRightOnly
            // 
            this.labelRightOnly.AutoSize = true;
            this.labelRightOnly.Dock = DockStyle.Fill;
            this.labelRightOnly.Location = new Point(3, 60);
            this.labelRightOnly.Name = "labelRightOnly";
            this.labelRightOnly.Size = new Size(140, 30);
            this.labelRightOnly.TabIndex = 4;
            this.labelRightOnly.Text = "Right Only Records:";
            this.labelRightOnly.TextAlign = ContentAlignment.MiddleLeft;
            
            // 
            // textBoxRightOnly
            // 
            this.textBoxRightOnly.Dock = DockStyle.Fill;
            this.textBoxRightOnly.Location = new Point(149, 63);
            this.textBoxRightOnly.Name = "textBoxRightOnly";
            this.textBoxRightOnly.ReadOnly = true;
            this.textBoxRightOnly.Size = new Size(336, 27);
            this.textBoxRightOnly.TabIndex = 5;
            
            // 
            // buttonExport
            // 
            this.buttonExport.Dock = DockStyle.Fill;
            this.buttonExport.Enabled = false;
            this.buttonExport.Location = new Point(3, 93);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new Size(140, 34);
            this.buttonExport.TabIndex = 6;
            this.buttonExport.Text = "Export Results";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new EventHandler(this.ButtonExport_Click);
            
            // 
            // progressBar1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.progressBar1, 2);
            this.progressBar1.Dock = DockStyle.Fill;
            this.progressBar1.Location = new Point(3, 633);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new Size(994, 64);
            this.progressBar1.TabIndex = 4;
            
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = Color.FromArgb(70, 130, 180);
            this.statusStrip1.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.statusStrip1.ImageScalingSize = new Size(20, 20);
            this.statusStrip1.Items.AddRange(new ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new Point(0, 774);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new Padding(1, 0, 16, 0);
            this.statusStrip1.Size = new Size(1200, 26);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.BackColor = Color.Transparent;
            this.toolStripStatusLabel1.ForeColor = Color.White;
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new Size(49, 20);
            this.toolStripStatusLabel1.Text = "Ready";
            
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel and CSV files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            this.openFileDialog1.Title = "Select Excel or CSV File";
            
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xlsx";
            this.saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            this.saveFileDialog1.Title = "Save Comparison Results with Original Data";
            
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new SizeF(8F, 20F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.FromArgb(240, 248, 255);
            this.ClientSize = new Size(1200, 800);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.statusStrip1);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.Icon = SystemIcons.Application;
            this.MinimumSize = new Size(1000, 700);
            this.Name = "MainForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Excel Lookup - Advanced File Comparison Tool v2.0";
            this.WindowState = FormWindowState.Maximized;
            
            this.tableLayoutPanel1.ResumeLayout(false);
            this.groupBoxLeft.ResumeLayout(false);
            this.tableLayoutPanelLeft.ResumeLayout(false);
            this.tableLayoutPanelLeft.PerformLayout();
            this.groupBoxRight.ResumeLayout(false);
            this.tableLayoutPanelRight.ResumeLayout(false);
            this.tableLayoutPanelRight.PerformLayout();
            this.groupBoxComparison.ResumeLayout(false);
            this.tableLayoutPanelComparison.ResumeLayout(false);
            this.tableLayoutPanelComparison.PerformLayout();
            this.groupBoxResults.ResumeLayout(false);
            this.tableLayoutPanelResults.ResumeLayout(false);
            this.tableLayoutPanelResults.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private GroupBox groupBoxLeft;
        private TableLayoutPanel tableLayoutPanelLeft;
        private Label labelLeftFile;
        private TextBox textBoxLeftFile;
        private Button buttonSelectLeftFile;
        private Label labelLeftSheet;
        private ComboBox comboBoxLeftSheet;
        private GroupBox groupBoxRight;
        private TableLayoutPanel tableLayoutPanelRight;
        private Label labelRightFile;
        private TextBox textBoxRightFile;
        private Button buttonSelectRightFile;
        private Label labelRightSheet;
        private ComboBox comboBoxRightSheet;
        private GroupBox groupBoxComparison;
        private TableLayoutPanel tableLayoutPanelComparison;
        private Label labelLeftColumns;
        private Label labelRightColumns;
        private CheckedListBox checkedListBoxLeftColumns;
        private CheckedListBox checkedListBoxRightColumns;
        private Button buttonAddMapping;
        private Label labelMappings;
        private ListBox listBoxMappings;
        private Button buttonRemoveMapping;
        private Button buttonCompare;
        private GroupBox groupBoxResults;
        private TableLayoutPanel tableLayoutPanelResults;
        private Label labelMatched;
        private TextBox textBoxMatched;
        private Label labelLeftOnly;
        private TextBox textBoxLeftOnly;
        private Label labelRightOnly;
        private TextBox textBoxRightOnly;
        private Button buttonExport;
        private ProgressBar progressBar1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private OpenFileDialog openFileDialog1;
        private SaveFileDialog saveFileDialog1;
    }
}