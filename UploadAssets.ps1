# Upload assets to existing release
param(
    [string]$Token = "$env:GITHUB_TOKEN"
)

# Get the release info to get upload URL
$headers = @{
    'Authorization' = "token $Token"
    'Accept' = 'application/vnd.github.v3+json'
}

try {
    $releaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/arehman626/ExcelLookupC/releases/tags/v1.0.0" -Headers $headers
    $uploadUrl = $releaseInfo.upload_url -replace '\{\?.*\}', ''
    
    Write-Host "🔗 Release URL: $($releaseInfo.html_url)" -ForegroundColor Cyan
    
    $assets = @(
        "ExcelLookupC-SingleFile-v1.0.0.zip",
        "ExcelLookupC-Standalone-v1.0.0.zip", 
        "ExcelLookupC-FrameworkDependent-v1.0.0.zip"
    )
    
    foreach ($assetFile in $assets) {
        if (Test-Path $assetFile) {
            Write-Host "📎 Uploading $assetFile..." -ForegroundColor Yellow
            
            $uploadHeaders = @{
                'Authorization' = "token $Token"
                'Content-Type' = 'application/zip'
            }
            
            $uri = "$uploadUrl" + "?name=$assetFile"
            $fileContent = [System.IO.File]::ReadAllBytes((Resolve-Path $assetFile))
            
            try {
                Invoke-RestMethod -Uri $uri -Method POST -Headers $uploadHeaders -Body $fileContent -ContentType "application/octet-stream"
                Write-Host "✅ Successfully uploaded $assetFile" -ForegroundColor Green
            } catch {
                Write-Host "⚠️  Error uploading $assetFile : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "❌ File not found: $assetFile" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "🎉 All done! Opening release page..." -ForegroundColor Green
    Start-Process $releaseInfo.html_url
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}