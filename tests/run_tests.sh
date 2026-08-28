#!/bin/bash
cd "$(dirname "$0")"
set -e # stop on first failure

# clean up
rm -f a.out

# build test suite
g++ -std=c++17 -I.. ../NonWindows.cpp tests.cpp

# run test suite
./a.out
