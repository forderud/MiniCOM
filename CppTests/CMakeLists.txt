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


set(CppTests_SOURCES
    Main.cpp
    AggregationTests.cpp
    BooleanTests.cpp
    ComClassTests.cpp
    ComErrorTests.cpp
    ComObjectTests.cpp
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
    # ComCreateTests uses OBJECT_ENTRY_AUTO and IUnknownFactory, which are the
    # non-Windows object-creation path. Windows uses ole32 and the registry.
    list(APPEND CppTests_SOURCES "../NonWindows.cpp" ComCreateTests.cpp)
endif()

add_executable(CppTests ${CppTests_SOURCES})

target_include_directories(CppTests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
)

target_link_libraries(CppTests PRIVATE TestInterfaces GTest::gtest)

# run test suite as post-build event
add_custom_command(TARGET CppTests POST_BUILD
                   COMMAND $<TARGET_FILE:CppTests>
                   WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
                   COMMENT "Running test suite")
