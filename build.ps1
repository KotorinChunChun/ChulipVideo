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
    app.py
