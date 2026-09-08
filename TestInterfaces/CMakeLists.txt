
set(IDL_INPUTS  "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl")
set(IDL_OUTPUTS "${CMAKE_CURRENT_BINARY_DIR}/Example.h")

add_library(TestInterfaces STATIC WrapperGen.cpp)

if(WIN32)
    target_sources(TestInterfaces PRIVATE "${IDL_INPUTS}")

    add_custom_command(TARGET TestInterfaces POST_BUILD
        COMMAND copy "$(IntDir)Example.h" .
        COMMAND py ..\\..\\TlhFilePatch.py "$(IntDir)Example.tlh" "${CMAKE_CURRENT_SOURCE_DIR}\\Example.tlh"
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
        COMMENT "Patching generated TLH file")

    target_include_directories(TestInterfaces PUBLIC ${CMAKE_CURRENT_BINARY_DIR}) # folder with generated headers
else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(TestInterfaces PRIVATE "${IDL_OUTPUTS}")
    target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_BINARY_DIR}") # folder with generated headers
endif()

target_include_directories(TestInterfaces PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}) # folder with generated headers
