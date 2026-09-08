
set(IDL_INPUTS  "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl")
set(IDL_OUTPUTS "${CMAKE_CURRENT_BINARY_DIR}/Example.h")

add_library(TestInterfaces STATIC Dummy.cpp)

if(WIN32)
    target_sources(TestInterfaces PRIVATE "${IDL_INPUTS}")

    # Add intermediate directory with MIDL output files to include path
    set(INT_DIR "${CMAKE_CURRENT_BINARY_DIR}/TestInterfaces.dir/${CMAKE_CFG_INTDIR}")
    target_include_directories(TestInterfaces PUBLIC "${INT_DIR}") # folder with generated headers
else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(TestInterfaces PRIVATE "${IDL_OUTPUTS}")
    target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_BINARY_DIR}") # folder with generated headers
endif()
