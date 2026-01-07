# GitHub Release Creation Script
param(
    [string]$Token = ""
)

if ([string]::IsNullOrEmpty($Token)) {
    Write-Host "🔑 GitHub Personal Access Token Required" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Go to: https://github.com/settings/tokens" -ForegroundColor Cyan
    Write-Host "2. Click 'Generate new token (classic)'" -ForegroundColor Cyan
    Write-Host "3. Select scopes: 'repo' (full repository access)" -ForegroundColor Cyan
    Write-Host "4. Copy the generated token" -ForegroundColor Cyan
    Write-Host ""
    Start-Process "https://github.com/settings/tokens"
    
    $Token = Read-Host "Enter your GitHub token"
}

if ([string]::IsNullOrEmpty($Token)) {
    Write-Host "❌ Token required. Exiting..." -ForegroundColor Red
    exit 1
}

Write-Host "🚀 Creating GitHub release..." -ForegroundColor Green

# Release data
$releaseData = @{
    tag_name = "v1.0.0"
    target_commitish = "main"
    name = "Excel Lookup v1.0.0"
    body = @"
# Excel Lookup v1.0.0 - Windows Forms Excel File Comparison Tool

## 🚀 Features
- **Dual Panel Interface**: Compare two Excel files side-by-side
- **Smart Sheet Selection**: Choose which sheets to compare from dropdown menus  
- **Column Selection**: Check boxes to select specific columns for comparison
- **Intelligent Comparison**: Finds matched, left-only, and right-only records
- **Sorted Results**: All comparison results automatically sorted for easy analysis
- **Professional Export**: Results exported to timestamped Excel file with metadata
- **Built-in Test Data**: Generate sample Excel files via Tools menu
- **Error Handling**: Comprehensive error handling with user-friendly messages
- **Async Operations**: Non-blocking UI with progress indicators

## 📦 Download Options

### 🎯 **Single-File Executable** (Recommended)
**File**: ``ExcelLookupC-SingleFile-v1.0.0.zip`` (66MB)
- Single ``ExcelLookupC.exe`` file
- **No dependencies required** - runs on any Windows system
- Extract and double-click to run immediately

### 📁 **Standalone Deployment**  
**File**: ``ExcelLookupC-Standalone-v1.0.0.zip`` (70MB)
- Complete deployment package with separate files
- **No dependencies required** - includes .NET runtime
- Professional deployment option

### ⚡ **Framework Dependent**
**File**: ``ExcelLookupC-FrameworkDependent-v1.0.0.zip`` (2MB)  
- Smallest download size
- **Requires .NET 8.0 Runtime** to be installed
- For development/technical users

## 🎮 How to Use
1. Download and extract your preferred version above
2. Run ``ExcelLookupC.exe`` 
3. Use **Tools → Generate Test Data** to create sample Excel files for testing
4. Select left and right Excel files using the file selection buttons
5. Choose sheets and comparison columns
6. Click **Compare** to analyze differences
7. Click **Export Results** to save timestamped comparison report

## 🛠️ Technical Details
- **Platform**: Windows (.NET 8.0)
- **Framework**: Windows Forms
- **Excel Library**: EPPlus 7.1.0
- **License**: MIT License
- **Source Code**: Full source code available in repository

## 📋 System Requirements
- **OS**: Windows 10/11
- **Memory**: 4GB RAM minimum
- **Disk Space**: 200MB free space
- **.NET Runtime**: Not required for single-file and standalone versions

## 🆘 Support
- **Issues**: Report bugs at [GitHub Issues](https://github.com/arehman626/ExcelLookupC/issues)
- **Documentation**: See [README.md](https://github.com/arehman626/ExcelLookupC/blob/main/README.md) for detailed usage
- **Test Guide**: See [TEST_PLAN.md](https://github.com/arehman626/ExcelLookupC/blob/main/TEST_PLAN.md)

---
**Built with ❤️ using C# Windows Forms and EPPlus**
"@
    draft = $false
    prerelease = $false
} | ConvertTo-Json -Depth 10

# Headers for GitHub API
$headers = @{
    'Authorization' = "token $Token"
    'Accept' = 'application/vnd.github.v3+json'
    'Content-Type' = 'application/json'
}

try {
    # Create the release
    Write-Host "📝 Creating release..." -ForegroundColor Yellow
    $releaseResponse = Invoke-RestMethod -Uri "https://api.github.com/repos/arehman626/ExcelLookupC/releases" -Method POST -Headers $headers -Body $releaseData
    
    Write-Host "✅ Release created successfully!" -ForegroundColor Green
    Write-Host "Release URL: $($releaseResponse.html_url)" -ForegroundColor Cyan
    
    # Upload assets
    $uploadUrl = $releaseResponse.upload_url -replace '\{\?.*\}', ''
    
    $assets = @(
        @{file="ExcelLookupC-SingleFile-v1.0.0.zip"; name="ExcelLookupC-SingleFile-v1.0.0.zip"}
        @{file="ExcelLookupC-Standalone-v1.0.0.zip"; name="ExcelLookupC-Standalone-v1.0.0.zip"}
        @{file="ExcelLookupC-FrameworkDependent-v1.0.0.zip"; name="ExcelLookupC-FrameworkDependent-v1.0.0.zip"}
    )
    
    foreach ($asset in $assets) {
        if (Test-Path $asset.file) {
            Write-Host "📎 Uploading $($asset.name)..." -ForegroundColor Yellow
            
            $uploadHeaders = @{
                'Authorization' = "token $Token"
                'Content-Type' = 'application/zip'
            }
            
            $fileBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $asset.file))
            
            $assetUploadUrl = "$uploadUrl" + "?name=$($asset.name)"
            $uploadResponse = Invoke-RestMethod -Uri $assetUploadUrl -Method POST -Headers $uploadHeaders -Body $fileBytes
            
            Write-Host "✅ Uploaded $($asset.name)" -ForegroundColor Green
        } else {
            Write-Host "⚠️  File not found: $($asset.file)" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "🎉 Release v1.0.0 created successfully!" -ForegroundColor Green
    Write-Host "🔗 View release: $($releaseResponse.html_url)" -ForegroundColor Cyan
    
    # Open the release page
    Start-Process $releaseResponse.html_url
    
} catch {
    Write-Host "❌ Error creating release: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.Exception.Response)" -ForegroundColor Red
    
    if ($_.Exception.Response.StatusCode -eq 401) {
        Write-Host "🔑 Authentication failed. Please check your token has 'repo' permissions." -ForegroundColor Yellow
    }
}