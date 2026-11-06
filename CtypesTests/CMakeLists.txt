# Windows builds a binding from the type library that MIDL emits, through
# comtypes, so there is nothing to generate and nothing to test there. It
# generates TestInterfaces.winmd instead, which the bindings here come from.
if(WIN32)
    return()
endif()

# Python bindings, generated from the .winmd file that the Windows build
# describes the IDL files in, and that is committed. Reading it takes dnfile.
find_package(Python COMPONENTS Interpreter REQUIRED)
execute_process(COMMAND ${Python_EXECUTABLE} -c "import dnfile" RESULT_VARIABLE NO_DNFILE OUTPUT_QUIET ERROR_QUIET)
if(NO_DNFILE)
    message(WARNING "The Python binding test is skipped, as its bindings are generated from TestInterfaces.winmd, which takes dnfile to read: ${Python_EXECUTABLE} -m pip install dnfile")
    return()
endif()

# ExampleBindings.py and RegistryBindings.py, which imports the former, as Registry.idl
# imports Example.idl
generate_python_bindings(ExampleBindings "${CMAKE_SOURCE_DIR}/TestInterfaces/TestInterfaces.winmd" Example Registry)
