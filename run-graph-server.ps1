# PowerShell script to run the MS Graph MCP server

# Get the current directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Function to check if Node.js is installed
function Check-NodeJs {
    try {
        $nodeVersion = node -v
        Write-Host "Node.js is installed: $nodeVersion"
        return $true
    }
    catch {
        Write-Host "Node.js is not installed or not in PATH"
        return $false
    }
}

# Function to check if npm is installed
function Check-Npm {
    try {
        $npmVersion = npm -v
        Write-Host "npm is installed: $npmVersion"
        return $true
    }
    catch {
        Write-Host "npm is not installed or not in PATH"
        return $false
    }
}

# Check if Node.js and npm are installed
$nodeInstalled = Check-NodeJs
$npmInstalled = Check-Npm

if (-not $nodeInstalled -or -not $npmInstalled) {
    Write-Host "Please install Node.js and npm before running this script."
    Write-Host "Download from: https://nodejs.org/"
    exit 1
}

# Check if node_modules exists, if not run npm install
if (-not (Test-Path "$scriptPath\node_modules")) {
    Write-Host "Installing dependencies..."
    Set-Location $scriptPath
    npm install
}

# Check if build directory exists, if not run build
if (-not (Test-Path "$scriptPath\build")) {
    Write-Host "Building project..."
    Set-Location $scriptPath
    npm run build
}

# Run the MCP server
Write-Host "Starting MS Graph MCP Server..."
Set-Location $scriptPath
npm start

# Keep the window open if there's an error
if ($LASTEXITCODE -ne 0) {
    Write-Host "An error occurred. Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
