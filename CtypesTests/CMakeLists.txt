# Windows builds a binding from the type library that MIDL emits, through
# comtypes, so there is nothing to generate and nothing to test there. It
# generates TestInterfaces.winmd instead, which the bindings here come from.
if(WIN32)
    return()
endif()

# COM server for the test to activate, in a shared library so that it can be
# loaded from outside C++
add_library(ExampleLib SHARED ExampleLib.cpp "../NonWindows.cpp")

target_include_directories(ExampleLib PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/..")

target_link_libraries(ExampleLib PRIVATE TestInterfaces)


# Python bindings, generated from the .winmd file that the Windows build
# describes the IDL files in, and that is committed. Reading it takes dnfile.
find_package(Python COMPONENTS Interpreter REQUIRED)
execute_process(COMMAND ${Python_EXECUTABLE} -c "import dnfile" RESULT_VARIABLE NO_DNFILE OUTPUT_QUIET ERROR_QUIET)
if(NO_DNFILE)
    message(WARNING "The Python binding test is skipped, as its bindings are generated from TestInterfaces.winmd, which takes dnfile to read: ${Python_EXECUTABLE} -m pip install dnfile")
    return()
endif()

# a namespace per IDL file. Registry.idl imports Example.idl, so its bindings
# import the ones for it as ExampleBindings
set(WINMD "${CMAKE_SOURCE_DIR}/TestInterfaces/TestInterfaces.winmd")
# the IDL files too, so that a change to one is warned about as soon as it is
# made, until the .winmd is regenerated on Windows
set(WINMD_SOURCES "${WINMD}.sha256" "${CMAKE_SOURCE_DIR}/TestInterfaces/Example.idl" "${CMAKE_SOURCE_DIR}/TestInterfaces/Registry.idl")
set(PY_BINDINGS "")
foreach(NAME Example Registry)
    set(PY_FILE "${CMAKE_CURRENT_BINARY_DIR}/${NAME}Bindings.py")
    add_custom_command(
        OUTPUT  ${PY_FILE}
        COMMAND ${Python_EXECUTABLE} "${CMAKE_CURRENT_SOURCE_DIR}/winmd2py.py" ${WINMD} ${NAME} ${PY_FILE}
        DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/winmd2py.py" ${WINMD} ${WINMD_SOURCES}
        COMMENT "Generating Python bindings for ${NAME}..."
    )
    list(APPEND PY_BINDINGS ${PY_FILE})
endforeach()
add_custom_target(ExampleBindings DEPENDS ${PY_BINDINGS})
add_dependencies(ExampleLib ExampleBindings)


# run python test script as post-build event
add_custom_command(TARGET ExampleLib POST_BUILD
    COMMAND PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR} ${Python_EXECUTABLE} tests.py $<TARGET_FILE:ExampleLib> ${WINMD}
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    COMMENT "Running python tests")
