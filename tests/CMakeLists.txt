cmake_minimum_required(VERSION 3.16)
project(MiniCOM_tests LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)

# Require GoogleTest:
# * Fedora: "dnf install gtest-devel"
# * MacOS: "brew install googletest"
find_package(GTest REQUIRED)

# compatibility function for parsing IDL files on non-Windows platforms
function(process_idl_files INPUT_FILES OUTPUTS_FILES)
    find_package(Python COMPONENTS Interpreter REQUIRED)

    add_custom_command(
        OUTPUT  ${OUTPUTS_FILES}
        COMMAND ${Python_EXECUTABLE} "${CMAKE_CURRENT_SOURCE_DIR}/../IdlParse.py" ${INPUT_FILES}
        DEPENDS ${INPUT_FILES}
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
        COMMENT "Processing IDL files..."
    )
endfunction()


set(TEST_SOURCES
    Main.cpp
    AggregationTests.cpp
    BooleanTests.cpp
    ComClassTests.cpp
    ComErrorTests.cpp
    ComObjectTests.cpp
    ComPtrTests.cpp
    DynamicCastTests.cpp
    SafeArrayTests.cpp
    SharedRefTests.cpp
    StringTests.cpp
    )
if(NOT WIN32)
    # ComCreateTests uses OBJECT_ENTRY_AUTO and IUnknownFactory, which are the
    # non-Windows object-creation path. Windows uses ole32 and the registry.
    list(APPEND TEST_SOURCES "../NonWindows.cpp" ComCreateTests.cpp)
endif()

set(IDL_INPUTS  "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl")
set(IDL_OUTPUTS "${CMAKE_CURRENT_BINARY_DIR}/Example.h")


add_definitions(-DUNICODE -D_UNICODE)

add_executable(MiniCOM_tests ${TEST_SOURCES})

target_include_directories(MiniCOM_tests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
    "${CMAKE_CURRENT_BINARY_DIR}"
)

if(NOT WIN32)
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(MiniCOM_tests PRIVATE "${IDL_OUTPUTS}")
endif()

target_link_libraries(MiniCOM_tests PRIVATE GTest::gtest)

# run test suite as post-build event
add_custom_command(TARGET MiniCOM_tests POST_BUILD
                   COMMAND $<TARGET_FILE:MiniCOM_tests>
                   WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
                   COMMENT "Running test suite")
