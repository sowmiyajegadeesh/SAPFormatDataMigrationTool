#!/bin/bash
# Bash script to set up Python environment, install dependencies, and run update_excel.py

# Check if .venv directory exists
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "Failed to create virtual environment. Ensure Python3 is installed."
        exit 1
    fi
fi

# Activate the virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate
if [ $? -ne 0 ]; then
    echo "Failed to activate virtual environment."
    exit 1
fi

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "Failed to install requirements."
    exit 1
fi

# Run the Python script
echo "Running update_excel.py..."
python3 update_excel.py
if [ $? -ne 0 ]; then
    echo "Script execution failed."
    exit 1
fi

echo "Done."