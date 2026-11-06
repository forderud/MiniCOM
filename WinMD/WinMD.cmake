# generate_winmd(<target> <file.winmd> [NAMESPACE <namespace>] <file.idl>...)
#
# Generates a .winmd file from IDL files, with one namespace per IDL file, named after it,
# and placed under <namespace> when given. Requires the .NET 10 SDK and the C++ build
# tools, so it only runs on Windows.
set(WINMD_PROJECT "${CMAKE_CURRENT_LIST_DIR}/WinMD.proj")

function(generate_winmd TARGET WINMD)
    cmake_parse_arguments(PARSE_ARGV 2 ARG "" "NAMESPACE" "")
    execute_process(COMMAND dotnet --list-sdks OUTPUT_VARIABLE DOTNET_SDKS ERROR_QUIET)
    if(NOT DOTNET_SDKS MATCHES "(^|\n)[1-9][0-9]+\\.")
        message(FATAL_ERROR "Unable to find .NET 10 SDK that's required for .winmd generation.")
    endif()

    get_filename_component(WINMD "${WINMD}" ABSOLUTE)
    set(IDL_FILES "")
    foreach(IDL_FILE ${ARG_UNPARSED_ARGUMENTS})
        get_filename_component(IDL_FILE "${IDL_FILE}" ABSOLUTE)
        list(APPEND IDL_FILES "${IDL_FILE}")
    endforeach()
    # the properties in a response file, as both the command line and MSBuild's -p: would
    # split the list of IDL files on its ";"
    set(OBJ_DIR "${CMAKE_CURRENT_BINARY_DIR}/${TARGET}")
    string(REPLACE ";" "%3B" IDL_LIST "${IDL_FILES}")
    set(RSP "${CMAKE_CURRENT_BINARY_DIR}/${TARGET}.rsp")
    file(GENERATE OUTPUT "${RSP}" CONTENT
"\"-p:IdlFiles=${IDL_LIST}\"
\"-p:OutputWinmd=${WINMD}\"
\"-p:WinmdNamespace=${ARG_NAMESPACE}\"
\"-p:ObjDir=${OBJ_DIR}\"
\"-p:BaseIntermediateOutputPath=${OBJ_DIR}/msbuild/\"
\"-p:BaseOutputPath=${OBJ_DIR}/bin/\"
")
    add_custom_command(
        OUTPUT  "${WINMD}"
        COMMAND ${CMAKE_COMMAND} -E rm -rf "${OBJ_DIR}" # https://github.com/microsoft/win32metadata/issues/2395
        COMMAND dotnet msbuild "${WINMD_PROJECT}" -restore -nologo -v:minimal "@${RSP}"
        DEPENDS ${IDL_FILES} "${WINMD_PROJECT}" "${RSP}"
        COMMENT "Generating ${WINMD}"
        VERBATIM)
    add_custom_target(${TARGET} ALL DEPENDS "${WINMD}")
endfunction()


# generate_python_bindings(<target> <file.winmd> <namespace>...)
#
# Generates ctypes bindings for namespaces of a .winmd file into <build dir>/<target>: one
# <name>Bindings.py module per namespace, named after its last part, so that MyApp.Example
# becomes ExampleBindings, and the comruntime.py they all use. Reading the .winmd takes
# dnfile.
set(WINMD_TO_PYTHON "${CMAKE_CURRENT_LIST_DIR}/winmd2py.py")
set(WINMD_COMRUNTIME "${CMAKE_CURRENT_LIST_DIR}/comruntime.py")

function(generate_python_bindings TARGET WINMD)
    find_package(Python COMPONENTS Interpreter REQUIRED)
    get_filename_component(WINMD "${WINMD}" ABSOLUTE)
    set(OUT_DIR "${CMAKE_CURRENT_BINARY_DIR}/${TARGET}")

    add_custom_command(
        OUTPUT  "${OUT_DIR}/comruntime.py"
        COMMAND ${CMAKE_COMMAND} -E copy "${WINMD_COMRUNTIME}" "${OUT_DIR}/comruntime.py"
        DEPENDS "${WINMD_COMRUNTIME}")
    set(OUTPUTS "${OUT_DIR}/comruntime.py")
    foreach(NAMESPACE ${ARGN})
        string(REGEX REPLACE "^.*\\." "" NAME "${NAMESPACE}")
        set(PY_FILE "${OUT_DIR}/${NAME}Bindings.py")
        add_custom_command(
            OUTPUT  "${PY_FILE}"
            COMMAND ${Python_EXECUTABLE} "${WINMD_TO_PYTHON}" "${WINMD}" "${NAMESPACE}" "${PY_FILE}"
            DEPENDS "${WINMD_TO_PYTHON}" "${WINMD}"
            COMMENT "Generating Python bindings for ${NAMESPACE}"
            VERBATIM)
        list(APPEND OUTPUTS "${PY_FILE}")
    endforeach()
    add_custom_target(${TARGET} ALL DEPENDS ${OUTPUTS})
endfunction()
