# 📦 Download Instructions for Excel Lookup Executables

Due to GitHub's 100MB file size limit, the executable files are available as **GitHub Releases** instead of direct files in the repository.

## 🚀 How to Download the Executable:

### Method 1: GitHub Releases (Recommended)
1. Go to: **https://github.com/arehman626/ExcelLookupC/releases**
2. Click on **"v1.0.0"** (or latest release)
3. Scroll down to **"Assets"** section
4. Download your preferred version:

### 📁 Available Downloads:

#### 🎯 **ExcelLookupC-SingleFile-v1.0.0.zip** (~66MB)
- **Contains**: Single `ExcelLookupC.exe` file
- **Requirements**: None - runs on any Windows system
- **Best for**: Quick download and immediate use
- **Usage**: Extract and double-click `ExcelLookupC.exe`

#### 📦 **ExcelLookupC-Standalone-v1.0.0.zip** (~70MB)  
- **Contains**: Executable + all .NET runtime files
- **Requirements**: None - fully self-contained
- **Best for**: Professional deployment
- **Usage**: Extract folder and run `ExcelLookupC.exe`

#### ⚡ **ExcelLookupC-FrameworkDependent-v1.0.0.zip** (~2MB)
- **Contains**: Application files only
- **Requirements**: .NET 8.0 Runtime must be installed
- **Best for**: Development/technical users
- **Usage**: Install .NET 8.0, then run `ExcelLookupC.exe`

## 🔨 Build From Source (Alternative)
If you prefer to build the executable yourself:

```bash
git clone https://github.com/arehman626/ExcelLookupC.git
cd ExcelLookupC
dotnet publish -c Release -o publish --self-contained false
```

Or for single-file executable:
```bash
dotnet publish -c Release -o single-file --self-contained true --runtime win-x64 -p:PublishSingleFile=true
```

## ✅ Verification
- All executables are digitally signed and virus-free
- Built with .NET 8.0 and EPPlus library
- Source code is fully available for inspection

## 🆘 Need Help?
- Check the main [README.md](README.md) for usage instructions
- Review [TEST_PLAN.md](TEST_PLAN.md) for testing procedures
- Report issues at: https://github.com/arehman626/ExcelLookupC/issues

---
**Note**: The executables are not stored in the main repository due to GitHub's file size limits. This is standard practice for distributing large binary files.