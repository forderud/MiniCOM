cmake_minimum_required(VERSION 3.16)
project(MiniCOM_tests LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)

# Require GoogleTest:
# * Fedora: "dnf install gtest-devel"
# * MacOS: "brew install googletest"
find_package(GTest REQUIRED)
include(GoogleTest)

set(TEST_SOURCES
    Main.cpp
    AggregationTests.cpp
    BooleanTests.cpp
    ComClassTests.cpp
    ComErrorTests.cpp
    ComPtrTests.cpp
    DynamicCastTests.cpp
    SafeArrayTests.cpp
    SharedRefTests.cpp
    StringTests.cpp
    )
if(NOT WIN32)
    list(APPEND TEST_SOURCES "../NonWindows.cpp")
endif()

add_definitions(-DUNICODE -D_UNICODE)

add_executable(MiniCOM_tests ${TEST_SOURCES})

target_include_directories(MiniCOM_tests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
)

target_link_libraries(MiniCOM_tests PRIVATE GTest::gtest)

enable_testing()
gtest_discover_tests(MiniCOM_tests WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}")
