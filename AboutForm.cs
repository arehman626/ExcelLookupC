using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelLookupC
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            LoadApplicationInfo();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.AutoScaleDimensions = new SizeF(8F, 20F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(500, 400);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutForm";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "About Excel Lookup";
            this.BackColor = Color.White;

            // Main panel with gradient background
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 248, 255) // Light blue background
            };
            this.Controls.Add(mainPanel);

            // Header panel
            var headerPanel = new Panel
            {
                Height = 80,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(70, 130, 180) // Steel blue
            };
            mainPanel.Controls.Add(headerPanel);

            // Application title
            var titleLabel = new Label
            {
                Text = "Excel Lookup",
                Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = false,
                Size = new Size(480, 40),
                Location = new Point(10, 10),
                TextAlign = ContentAlignment.MiddleLeft
            };
            headerPanel.Controls.Add(titleLabel);

            // Subtitle
            var subtitleLabel = new Label
            {
                Text = "Advanced File Comparison Tool",
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                AutoSize = false,
                Size = new Size(480, 25),
                Location = new Point(10, 45),
                TextAlign = ContentAlignment.MiddleLeft
            };
            headerPanel.Controls.Add(subtitleLabel);

            // Content panel
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                BackColor = Color.Transparent
            };
            mainPanel.Controls.Add(contentPanel);

            // Description text
            var descriptionLabel = new Label
            {
                Text = @"Excel Lookup is a powerful file comparison application designed to help you analyze and compare data between Excel spreadsheets and CSV files.

Key Features:
• Dual-panel interface for easy file selection
• Support for Excel (.xlsx, .xls) and CSV files
• Flexible column mapping for different file structures
• Advanced comparison with matched, left-only, and right-only records
• Export results with original data preservation
• Built-in test data generation

Perfect for data analysts, accountants, and anyone who needs to compare structured data files efficiently.",
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60),
                BackColor = Color.Transparent,
                AutoSize = false,
                Size = new Size(460, 200),
                Location = new Point(0, 0),
                TextAlign = ContentAlignment.TopLeft
            };
            contentPanel.Controls.Add(descriptionLabel);

            // Version and copyright info
            var versionLabel = new Label
            {
                Text = "Version 2.0.0",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(70, 130, 180),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 220)
            };
            contentPanel.Controls.Add(versionLabel);

            var copyrightLabel = new Label
            {
                Text = "© 2026 Excel Lookup Application",
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100),
                BackColor = Color.Transparent,
                AutoSize = true,
                Location = new Point(0, 245)
            };
            contentPanel.Controls.Add(copyrightLabel);

            // OK button
            var okButton = new Button
            {
                Text = "OK",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Size = new Size(80, 30),
                Location = new Point(210, 280),
                BackColor = Color.FromArgb(70, 130, 180),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false
            };
            okButton.FlatAppearance.BorderSize = 0;
            okButton.Click += (s, e) => this.Close();
            contentPanel.Controls.Add(okButton);

            this.ResumeLayout(false);
        }

        private void LoadApplicationInfo()
        {
            // This method can be extended to load dynamic information
        }
    }
}