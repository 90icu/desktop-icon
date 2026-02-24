$ErrorActionPreference = "Stop"

Write-Host "Starting build process..."

# Extract icon if it doesn't exist
if (-not (Test-Path "app.ico")) {
    Write-Host "Extracting icon from ttkbootstrap..."
    python extract_icon.py
}

# Clean previous build artifacts
# Use try-catch or -ErrorAction SilentlyContinue for deletion to handle open files
try {
    if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path "dist") { Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path "*.spec") { Remove-Item -Path "*.spec" -Force -ErrorAction SilentlyContinue }
} catch {
    Write-Host "Warning: Could not clean some previous artifacts. They might be in use."
}

# Run PyInstaller
# --noconsole: Hide terminal window
# --onefile: Bundle everything into a single EXE
# --name: Name of the output file
# --hidden-import: Ensure ttkbootstrap is included
# --icon: Set the executable icon
# --add-data: Bundle app.ico so it can be used at runtime if needed (though --onefile makes this tricky for direct access, but --icon handles the EXE file)
pyinstaller --noconsole --onefile --name "DesktopManager_Temp" --clean --collect-all ttkbootstrap --collect-all pystray --icon "app.ico" --add-data "app.ico;." main_gui.py

if ($LASTEXITCODE -eq 0) {
    python rename_dist.py
} else {
    Write-Host "Build failed!"
}
