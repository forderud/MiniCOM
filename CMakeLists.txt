cmake_minimum_required(VERSION 3.16)
project(MiniCOM LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)

if(MSVC)
    # warning level 4
    add_compile_options(/W4)
else()
    # -Wall and -Wextra corresponds roughly to /W4 used for MSVC,
    # see https://clang.llvm.org/docs/UsersManual.html#id9
    add_compile_options(-Wall)
    add_compile_options(-Wextra)
endif()

add_definitions(-DUNICODE -D_UNICODE)

if(NOT WIN32)
    # need to capture current dir outside of the function
    set(CURRENT_DIR ${CMAKE_CURRENT_LIST_DIR})

    # compatibility function for parsing IDL files on non-Windows platforms
    function(process_idl_files INPUT_FILES OUTPUTS_FILES)
        find_package(Python COMPONENTS Interpreter REQUIRED)

        add_custom_command(
            OUTPUT  ${OUTPUTS_FILES}
            COMMAND ${Python_EXECUTABLE} "${CURRENT_DIR}/IdlParse.py" ${INPUT_FILES}
            DEPENDS ${INPUT_FILES}
            WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
            COMMENT "Processing IDL files..."
        )
    endfunction()
endif()

add_subdirectory(TestInterfaces)
add_subdirectory(CppTests)
add_subdirectory(PyTests)
add_subdirectory(CtypesTests)
