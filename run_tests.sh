#!/bin/bash
cd "$(dirname "$0")"
set -e # stop on first failure

# Clean up
rm -rf build

# Generate project files
cmake -S . -B build -DPython_EXECUTABLE="$(which python3)"

# Build & run tests
cmake --build build
