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

# run python test script as post-build event
if(WIN32)
    add_custom_command(TARGET PyTests POST_BUILD
        COMMAND set PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR}\\$<IF:$<CONFIG:Debug>,Debug,Release>
        COMMAND py tests.py
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Running python tests")
else()
    add_custom_command(TARGET PyTests POST_BUILD
        COMMAND PYTHONPATH=${CMAKE_CURRENT_BINARY_DIR} ${Python_EXECUTABLE} tests.py
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Running python tests")
endif()
