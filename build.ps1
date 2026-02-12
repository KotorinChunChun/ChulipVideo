python -m PyInstaller --noconsole --onefile --clean --noconfirm `
    --collect-all cv2 --collect-all PIL `
    --hidden-import=PIL._tkinter_finder `
    --hidden-import=ChunapTool `
    --hidden-import=CapchunScreen `
    --hidden-import=ChulipVideo `
    --add-data "README_shortcuts.md;." `
    --add-data "README.md;." `
    --add-data "ChulipVideo.png;." `
    --icon="ChulipVideo.png" `
    --name="ChulipVideo" `
    --paths="src" `
    src/app.py

# ビルド後のファイルコピー
Write-Host "Copying additional files to dist..."
$dist = "dist"

# shortcutsフォルダ
if (Test-Path "shortcuts") {
    Copy-Item -Path "shortcuts" -Destination $dist -Recurse -Force
}

# batファイルすべて
if (Test-Path "build_tools\*.bat") {
    Copy-Item -Path "build_tools\*.bat" -Destination $dist -Force
}

Write-Host "Build and copy complete."
