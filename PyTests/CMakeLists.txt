find_package(Python REQUIRED COMPONENTS Interpreter Development.Module) # Interpreter needed for Python_SITELIB below

set(pybind11_DIR ${Python_SITELIB}/pybind11/share/cmake/pybind11) # "sudo pip install pybind11"
find_package(pybind11 CONFIG REQUIRED)


set(PyTests_SOURCES
    Main.cpp
    tests.py
    )
if(NOT WIN32)
    list(APPEND PyTests_SOURCES "../NonWindows.cpp")
endif()

python_add_library(PyTests MODULE ${PyTests_SOURCES})

target_include_directories(PyTests PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/.."
)

target_link_libraries(PyTests PUBLIC
        TestInterfaces
        Python::Module
        pybind11::headers
        )

# run python test script after build, and whenever tests.py changes
set(TEST_STAMP "${CMAKE_CURRENT_BINARY_DIR}/tests.stamp")
if(WIN32)
    add_custom_command(OUTPUT ${TEST_STAMP}
        COMMAND set PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR}\\$<IF:$<CONFIG:Debug>,Debug,Release>
        COMMAND py tests.py
        COMMAND ${CMAKE_COMMAND} -E touch ${TEST_STAMP}
        DEPENDS tests.py PyTests
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Running python tests")
else()
    add_custom_command(OUTPUT ${TEST_STAMP}
        COMMAND PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR} ${Python_EXECUTABLE} tests.py
        COMMAND ${CMAKE_COMMAND} -E touch ${TEST_STAMP}
        DEPENDS tests.py PyTests
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Running python tests")
endif()
add_custom_target(PyTests_run ALL DEPENDS ${TEST_STAMP})
add_dependencies(PyTests_run PyTests) # ensure PyTests.pyd exists before the script runs
