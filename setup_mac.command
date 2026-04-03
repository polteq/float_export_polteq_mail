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

APP_PATH="$HOME/Desktop/Process Timesheet.app"
rm -rf "$APP_PATH"

# Create a temporary AppleScript source file
cat > "tmp_applet.applescript" << EOF
on open dropped_files
    repeat with the_file in dropped_files
        set posix_path to POSIX path of the_file
        tell application "Terminal"
            activate
            do script "cd '$PWD' && source venv/bin/activate && python3 shareable_processor.py '" & posix_path & "'; exit"
        end tell
    end repeat
end open

on run
    tell application "Terminal"
        activate
        do script "cd '$PWD' && source venv/bin/activate && python3 shareable_processor.py; exit"
    end tell
end run
EOF

# Compile the AppleScript into a real application
osacompile -o "$APP_PATH" "tmp_applet.applescript"
rm "tmp_applet.applescript"

# Add a basic icon hint by setting the CFBundleIdentifier (optional but cleaner)
/usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier com.polteq.timesheetprocessor" "$APP_PATH/Contents/Info.plist"

echo ""
echo "Setup Complete!"
echo "You can now drag and drop your Timesheet CSV files onto the 'Process Timesheet' application on your Desktop."
echo "Note: The first time you drop a file on it, you might need to Right-Click -> Open to allow execution."
