if(WIN32)
    include(FetchContent)

    FetchContent_Declare(googletest
      GIT_REPOSITORY https://github.com/google/googletest.git
      GIT_TAG        v1.18.0
    )

    FetchContent_MakeAvailable(googletest)
else()
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
endif()


set(TEST_SOURCES
    Main.cpp
    AggregationTests.cpp
    BooleanTests.cpp
    ComClassTests.cpp
    ComErrorTests.cpp
    ComPtrTests.cpp
    DynamicCastTests.cpp
    ExampleTests.cpp
    ExampleVtableTests.hpp
    ExampleVtableTests.cpp
    SafeArrayTests.cpp
    SharedRefTests.cpp
    StringTests.cpp
    )
if(NOT WIN32)
    list(APPEND TEST_SOURCES "../NonWindows.cpp")
endif()

set(IDL_INPUTS  "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl")
set(IDL_OUTPUTS "${CMAKE_CURRENT_BINARY_DIR}/Example.h")

add_executable(MiniCOM_tests ${TEST_SOURCES})

target_include_directories(MiniCOM_tests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
    "${CMAKE_CURRENT_BINARY_DIR}"
)

if(WIN32)
    target_sources(MiniCOM_tests PRIVATE "${IDL_INPUTS}")

    # Add intermediate directory with MIDL output files to include path
    set(INT_DIR "${CMAKE_CURRENT_BINARY_DIR}/MiniCOM_tests.dir/${CMAKE_CFG_INTDIR}")
    target_include_directories(MiniCOM_tests PRIVATE "${INT_DIR}")
else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(MiniCOM_tests PRIVATE "${IDL_OUTPUTS}")
endif()

target_link_libraries(MiniCOM_tests PRIVATE GTest::gtest)

# run test suite as post-build event
add_custom_command(TARGET MiniCOM_tests POST_BUILD
                   COMMAND $<TARGET_FILE:MiniCOM_tests>
                   WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
                   COMMENT "Running test suite")
