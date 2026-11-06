
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

    # TestInterfaces.winmd, committed like Example.tlh. Regenerating it takes the .NET 10 SDK.
    execute_process(COMMAND dotnet --list-sdks OUTPUT_VARIABLE DOTNET_SDKS ERROR_QUIET)
    if(DOTNET_SDKS MATCHES "(^|\n)[1-9][0-9]+\\.")
        set(WINMD "${CMAKE_CURRENT_SOURCE_DIR}/TestInterfaces.winmd")
        file(GLOB WINMD_PROJECT LIST_DIRECTORIES false "${CMAKE_SOURCE_DIR}/WinMD/*")
        add_custom_command(
            OUTPUT  ${WINMD}
            COMMAND ${CMAKE_COMMAND} -E rm -rf obj bin # https://github.com/microsoft/win32metadata/issues/2395
            COMMAND dotnet build TestInterfaces.proj -nologo -v:minimal
            DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl" ${WINMD_PROJECT}
            WORKING_DIRECTORY "${CMAKE_SOURCE_DIR}/WinMD"
            COMMENT "Generating TestInterfaces.winmd")
        add_custom_target(TestInterfacesWinMD ALL DEPENDS ${WINMD})
    else()
        message(STATUS "No .NET 10 SDK, so TestInterfaces.winmd is not regenerated")
    endif()
else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(TestInterfaces PRIVATE "${IDL_OUTPUTS}")
    target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_BINARY_DIR}") # folder with generated headers
endif()

target_include_directories(TestInterfaces PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}) # folder with generated headers
