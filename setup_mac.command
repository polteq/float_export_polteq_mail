#!/bin/bash
cd "$(dirname "$0")"

echo "========================================="
echo "Polteq Timesheet Processor Setup (macOS)"
echo "========================================="

echo "Checking for Python..."
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install it (e.g., via Homebrew: brew install python3)."
    exit 1
fi

echo "Creating virtual environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

echo "Activating virtual environment and installing dependencies..."
source venv/bin/activate
pip install -r requirements.txt

echo ""
echo "Running initial configuration..."
python3 shareable_processor.py --setup

echo ""
echo "Creating Desktop Drag-and-Drop Application..."

APP_DIR="$HOME/Desktop/Process Timesheet.app"
mkdir -p "$APP_DIR/Contents/MacOS"

# Create Info.plist to make it a drop target
cat > "$APP_DIR/Contents/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>dropper</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.polteq.timesheetprocessor</string>
    <key>CFBundleName</key>
    <string>Process Timesheet</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleSignature</key>
    <string>????</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.10</string>
    <key>CFBundleDocumentTypes</key>
    <array>
        <dict>
            <key>CFBundleTypeExtensions</key>
            <array>
                <string>csv</string>
            </array>
            <key>CFBundleTypeRole</key>
            <string>Viewer</string>
        </dict>
    </array>
</dict>
</plist>
EOF

# Create the bash executable that receives the dropped file
cat > "$APP_DIR/Contents/MacOS/dropper" << EOF
#!/bin/bash
cd "$PWD"
source venv/bin/activate
# Argument 1 is the file dropped
if [ -n "\$1" ]; then
    python3 shareable_processor.py "\$1"
else
    # Fallback to run without args
    python3 shareable_processor.py
fi
EOF

chmod +x "$APP_DIR/Contents/MacOS/dropper"

echo ""
echo "Setup Complete!"
echo "You can now drag and drop your Timesheet CSV files onto the 'Process Timesheet' application on your Desktop."
echo "Note: The first time you drop a file on it, you might need to Right-Click -> Open to allow execution."
