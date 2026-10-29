
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

    # the interfaces as a .winmd file, the metadata format of win32metadata.
    # Generating it takes MIDL and the .NET SDK, so it is committed, the way
    # Example.tlh is, for other platforms to read.
    find_program(DOTNET_EXECUTABLE dotnet)
    execute_process(COMMAND py -c "import dnfile" RESULT_VARIABLE NO_DNFILE OUTPUT_QUIET ERROR_QUIET)
    if(NOT DOTNET_EXECUTABLE)
        message(STATUS "No .NET SDK found, so TestInterfaces.winmd is not regenerated")
    elseif(NO_DNFILE)
        message(STATUS "No dnfile for py, so TestInterfaces.winmd is not regenerated: py -m pip install dnfile")
    else()
        file(GLOB WINMD_PROJECT LIST_DIRECTORIES false "${CMAKE_SOURCE_DIR}/Winmd/*")
        # regenerated only when the IDL or the project changes, and from a
        # clean obj folder, since the generator's own incremental build does
        # not notice every change to the project
        add_custom_command(
            OUTPUT  "${CMAKE_CURRENT_SOURCE_DIR}/TestInterfaces.winmd"
            COMMAND ${CMAKE_COMMAND} -E rm -rf obj bin
            COMMAND ${DOTNET_EXECUTABLE} build TestInterfaces.proj -nologo -v:minimal
            COMMAND py WinmdRemoveTimeStamps.py "${CMAKE_CURRENT_SOURCE_DIR}/TestInterfaces.winmd"
            DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/Example.idl" ${WINMD_PROJECT}
            WORKING_DIRECTORY "${CMAKE_SOURCE_DIR}/Winmd"
            COMMENT "Describing the interfaces in TestInterfaces.winmd")
        add_custom_target(TestInterfacesWinmd ALL DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/TestInterfaces.winmd")
    endif()
else()
    process_idl_files("${IDL_INPUTS}" "${IDL_OUTPUTS}")
    target_sources(TestInterfaces PRIVATE "${IDL_OUTPUTS}")
    target_include_directories(TestInterfaces PUBLIC "${CMAKE_CURRENT_BINARY_DIR}") # folder with generated headers
endif()

target_include_directories(TestInterfaces PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}) # folder with generated headers
