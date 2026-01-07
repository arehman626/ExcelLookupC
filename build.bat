@echo off
echo Building Excel Lookup Application for deployment...

echo Cleaning previous builds...
dotnet clean --configuration Release

echo Restoring packages...
dotnet restore

echo Publishing application...
dotnet publish --configuration Release --output "./publish" --self-contained false --framework net8.0-windows

echo Build complete! Files are in the 'publish' folder.
echo To run the application, navigate to the publish folder and run ExcelLookupC.exe
pause