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

add_executable(MiniCOM_tests ${TEST_SOURCES})

target_include_directories(MiniCOM_tests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
)

target_link_libraries(MiniCOM_tests PRIVATE TestInterfaces GTest::gtest)

# run test suite as post-build event
add_custom_command(TARGET MiniCOM_tests POST_BUILD
                   COMMAND $<TARGET_FILE:MiniCOM_tests>
                   WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
                   COMMENT "Running test suite")
