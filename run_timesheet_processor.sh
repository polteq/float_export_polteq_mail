#!/bin/bash
# Timesheet Processor Runner Script

cd "$(dirname "$0")"

echo "=================================="
echo "Timesheet Processor"
echo "=================================="
echo ""

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    exit 1
fi

# Check if dependencies are installed
if ! python3 -c "import pandas" 2>/dev/null; then
    echo "Installing dependencies..."
    pip3 install -r requirements.txt
    echo ""
fi

# Run the processor
echo "Running timesheet processor..."
echo ""
python3 process_timesheet.py

echo ""
echo "Press any key to exit..."
read -n 1
