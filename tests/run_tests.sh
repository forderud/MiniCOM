#!/bin/bash
cd "$(dirname "$0")"
set -e # stop on first failure

# clean up
rm -f a.out

# build test suite
# This require GoogleTest ("gtest-devel" on Fedora)
g++ -std=c++17 -I. -I.. -lgtest ../NonWindows.cpp Main.cpp AggregationTests.cpp BooleanTests.cpp ComClassTests.cpp ComErrorTests.cpp ComPtrTests.cpp DynamicCastTests.cpp SafeArrayTests.cpp SharedRefTests.cpp StringTests.cpp

# run test suite
./a.out
