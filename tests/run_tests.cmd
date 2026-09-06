@echo off

:: Clean up
rmdir /s /q build

:: Generate project files
cmake -S . -B build

:: Open project in Visual Studio
cmake --open build
