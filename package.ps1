$date = Get-Date
$version = $date.ToString("yyyy-dd-M--HH-mm-ss")
$filename = "S3Spreadsheet-" + $version + ".zip"
cd .\S3Spreadsheet\src\S3Spreadsheet
dotnet lambda package ..\..\..\Packages\$filename --configuration Release -frun dotnet6 -farch arm64
cd ..\..\..