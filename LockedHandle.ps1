$TempFolder = "$env:TEMP\HandleTool"
$HandleExe = "$TempFolder\handle.exe"
$ZipFile = "$TempFolder\handle.zip"
$HandleUrl = "https://download.sysinternals.com/files/Handle.zip"

if (-not (Test-Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
}

if (-not (Test-Path $HandleExe)) {
    Write-Host "Downloading Handle.exe..."
    Invoke-WebRequest -Uri $HandleUrl -OutFile $ZipFile

    Write-Host "Extracting Handle.exe..."
    Expand-Archive -Path $ZipFile -DestinationPath $TempFolder -Force

    Remove-Item $ZipFile -Force
}

$TargetFile = Read-Host "Enter file path"

if (-not (Test-Path $TargetFile)) {
    Write-Host "Error: File does not exist. Please check the path and try again." -ForegroundColor Red
    Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
    Read-Host
    exit
}

# Capture handle.exe output
$ProcessInfo = & $HandleExe -accepteula $TargetFile 2>$null

$FilteredOutput = $ProcessInfo | Where-Object {$_ -notmatch "Nthandle|Copyright|Sysinternals"}

# Display only relevant output
if ($FilteredOutput) {
    $FilteredOutput | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "No process is locking the file."
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
Read-Host
