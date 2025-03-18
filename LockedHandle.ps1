# Define temp folder and handle.exe path
$TempFolder = "$env:TEMP\HandleTool"
$HandleExe = "$TempFolder\handle.exe"
$ZipFile = "$TempFolder\handle.zip"
$HandleUrl = "https://download.sysinternals.com/files/Handle.zip"

# Ensure temp folder exists
if (-not (Test-Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
}

# Download and extract Handle.exe if not already present
if (-not (Test-Path $HandleExe)) {
    Write-Host "Downloading Handle.exe..."
    Invoke-WebRequest -Uri $HandleUrl -OutFile $ZipFile

    # Extract the zip file
    Write-Host "Extracting Handle.exe..."
    Expand-Archive -Path $ZipFile -DestinationPath $TempFolder -Force

    # Clean up zip file
    Remove-Item $ZipFile -Force
}

# Ask the user for the file path
$TargetFile = Read-Host "Enter file path"

# Ensure the file exists
if (-not (Test-Path $TargetFile)) {
    Write-Host "Error: File does not exist. Please check the path and try again." -ForegroundColor Red
    Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
    Read-Host
    exit
}

# Capture handle.exe output
$ProcessInfo = & $HandleExe -accepteula $TargetFile 2>$null

# Remove header text from the output
$FilteredOutput = $ProcessInfo | Where-Object {$_ -notmatch "Nthandle|Copyright|Sysinternals"}

# Display only relevant output
if ($FilteredOutput) {
    $FilteredOutput | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "No process is locking the file."
}

# Keep window open
Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
Read-Host
