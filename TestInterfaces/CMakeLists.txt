
set(IDL_INPUTS  "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl")
set(IDL_OUTPUTS "${CMAKE_CURRENT_BINARY_DIR}/Example.h")

add_library(TestInterfaces STATIC WrapperGen.cpp)

target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}") # folder with committed TLH wrappers

if(WIN32)
    target_sources(TestInterfaces PRIVATE "${IDL_INPUTS}")

    # Add intermediate directory with MIDL output files to include path
    set(INT_DIR "${CMAKE_CURRENT_BINARY_DIR}/TestInterfaces.dir/${CMAKE_CFG_INTDIR}")
    target_include_directories(TestInterfaces PUBLIC "${INT_DIR}") # folder with generated headers

    # verify that the committed TLH/TLI wrappers match what the current IDL produces
    find_package(Python REQUIRED COMPONENTS Interpreter)
    set(REGEN_DIR "${CMAKE_CURRENT_BINARY_DIR}/regen")
    add_custom_command(TARGET TestInterfaces POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory "${REGEN_DIR}"
        COMMAND ${Python_EXECUTABLE} "${CMAKE_SOURCE_DIR}/TlhFilePatch.py" "${INT_DIR}/Example.tlh" "${REGEN_DIR}/ExampleWrap.tlh"
        COMMAND ${CMAKE_COMMAND} -DREGEN_DIR=${REGEN_DIR} -DSOURCE_DIR=${CMAKE_CURRENT_SOURCE_DIR} -P "${CMAKE_CURRENT_SOURCE_DIR}/CheckWrappers.cmake"
        COMMENT "Checking that the committed TLH wrappers are up to date")

else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(TestInterfaces PRIVATE "${IDL_OUTPUTS}")
    target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_BINARY_DIR}") # folder with generated headers
endif()
