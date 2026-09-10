# Windows builds a binding from the type library that MIDL emits, through
# comtypes, so there is nothing to generate and nothing to test there.
if(WIN32)
    return()
endif()

# COM server for the test to activate, in a shared library so that it can be
# loaded from outside C++
add_library(ExampleLib SHARED ExampleLib.cpp "../NonWindows.cpp")

target_include_directories(ExampleLib PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/..")

target_link_libraries(ExampleLib PRIVATE TestInterfaces)
