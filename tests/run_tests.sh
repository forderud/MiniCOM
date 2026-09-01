#!/bin/bash
cd "$(dirname "$0")"
set -e # stop on first failure

# Clean up
rm -rf build

# Generate project files
cmake -S . -B build

# Build & run tests
cmake --build build
cd build && ctest --output-on-failure
