# Excel Lookup Application - Copilot Instructions

## Project Overview
This is a C# Windows Forms application for comparing Excel files with the following features:
- Dual panel interface for selecting left and right Excel files
- Sheet selection from loaded Excel files
- Column selection for comparison
- Sorted comparison results
- Export results to timestamped Excel file with matched, left-only, and right-only records

## Development Guidelines
- Use Windows Forms for GUI
- Use EPPlus library for Excel file manipulation
- Implement proper error handling for file operations
- Follow C# coding conventions
- Include comprehensive testing
- Maintain clean separation of concerns with proper class structure

## Key Components
- MainForm: Primary GUI with file selection panels
- ExcelComparer: Core comparison logic
- ComparisonResult: Data model for results
- ExcelExporter: Handles result export functionality