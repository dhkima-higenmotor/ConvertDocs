#!/bin/bash
# This script runs the ConvertDocs.py application using the uv run command.

# Get the directory where the script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Change to the script's directory to ensure relative paths work correctly
cd "$SCRIPT_DIR"

# Execute the python script using uv
uv run ConvertDocs.py
