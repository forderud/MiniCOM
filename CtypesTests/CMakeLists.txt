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


# Python bindings, generated from the JSON description that IdlParse.py writes
# beside the header
find_package(Python COMPONENTS Interpreter REQUIRED)

# Registry.idl imports Example.idl, so its bindings read Example.json as well,
# and import the ones for it as ExampleBindings
set(JSON_FILES "${CMAKE_BINARY_DIR}/TestInterfaces/Example.json" "${CMAKE_BINARY_DIR}/TestInterfaces/Registry.json")
set(PY_BINDINGS "")
foreach(NAME Example Registry)
    set(PY_FILE "${CMAKE_CURRENT_BINARY_DIR}/${NAME}Bindings.py")
    add_custom_command(
        OUTPUT  ${PY_FILE}
        COMMAND ${Python_EXECUTABLE} "${CMAKE_CURRENT_SOURCE_DIR}/idl2py.py" "${CMAKE_BINARY_DIR}/TestInterfaces/${NAME}.json" ${PY_FILE}
        DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/idl2py.py" ${JSON_FILES}
        COMMENT "Generating Python bindings for ${NAME}..."
    )
    list(APPEND PY_BINDINGS ${PY_FILE})
endforeach()
add_custom_target(ExampleBindings DEPENDS ${PY_BINDINGS})
add_dependencies(ExampleBindings TestInterfaces) # for Example.json
add_dependencies(ExampleLib ExampleBindings)


# run python test script as post-build event
add_custom_command(TARGET ExampleLib POST_BUILD
    COMMAND PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR} ${Python_EXECUTABLE} tests.py $<TARGET_FILE:ExampleLib>
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    COMMENT "Running python tests")
