# PowerShell Script: Fix Poppler Dependency for pdf2image with Correct Folder Detection

Write-Host "Checking and fixing Poppler dependency for 'pdf2image'..." -ForegroundColor Cyan

# Define Poppler download URL and installation directory
$popplerZipUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"
$popplerBaseInstallDir = "C:\Poppler"
$popplerBinPath = ""

# Step 1: Check if Poppler is already installed
Write-Host "Checking for Poppler installation..."
if (Test-Path "$popplerBaseInstallDir") {
    $popplerBinPath = Get-ChildItem -Path "$popplerBaseInstallDir" -Recurse |
        Where-Object { $_.Name -eq "pdftoppm.exe" } |
        Select-Object -First 1 -ExpandProperty DirectoryName

    if ($popplerBinPath) {
        Write-Host "Poppler is already installed at $popplerBinPath." -ForegroundColor Green
    } else {
        Write-Host "Poppler installation exists, but 'pdftoppm.exe' is missing. Downloading fresh installation..."
    }
} else {
    Write-Host "Poppler is not installed. Proceeding with installation..."
}

# Step 2: Download and install Poppler if necessary
if (-not $popplerBinPath) {
    # Define paths for downloading and extracting Poppler
    $popplerZipPath = "$env:TEMP\Poppler.zip"

    # Download Poppler zip file
    Write-Host "Downloading Poppler from $popplerZipUrl..."
    Invoke-WebRequest -Uri $popplerZipUrl -OutFile $popplerZipPath -UseBasicParsing

    # Extract Poppler zip file to base install directory
    Write-Host "Extracting Poppler to $popplerBaseInstallDir..."
    Expand-Archive -Path $popplerZipPath -DestinationPath $popplerBaseInstallDir -Force

    # Check for the location of the pdftoppm.exe file
    $popplerBinPath = Get-ChildItem -Path "$popplerBaseInstallDir" -Recurse |
        Where-Object { $_.Name -eq "pdftoppm.exe" } |
        Select-Object -First 1 -ExpandProperty DirectoryName

    # Validate installation
    if (-NOT $popplerBinPath) {
        Write-Host "Poppler installation failed. Please download and install manually from $popplerZipUrl." -ForegroundColor Red
        Exit 1
    }

    Write-Host "Poppler installed successfully at $popplerBinPath." -ForegroundColor Green
}

# Step 3: Add Poppler to PATH
Write-Host "Adding Poppler to the system PATH..."
$path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
if ($path -notlike "*$popplerBinPath*") {
    [System.Environment]::SetEnvironmentVariable("Path", "$path;$popplerBinPath", [System.EnvironmentVariableTarget]::Machine)
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
    Write-Host "Poppler added to the system PATH." -ForegroundColor Green
} else {
    Write-Host "Poppler is already in the system PATH. Skipping PATH modification..." -ForegroundColor Yellow
}

# Step 4: Verify Poppler installation
Write-Host "Verifying Poppler installation..."
$popplerCheck = & pdftoppm -h 2>&1
if ($popplerCheck -match "pdftoppm") {
    Write-Host "Poppler is installed and functional!" -ForegroundColor Green
} else {
    Write-Host "Poppler verification failed! Ensure it is properly installed and in PATH." -ForegroundColor Red
    Exit 1
}

Write-Host "Poppler dependency fixed. You can now re-run the Python script!" -ForegroundColor Cyan