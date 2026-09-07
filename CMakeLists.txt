cmake_minimum_required(VERSION 3.16)
project(MiniCOM_tests LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)

add_definitions(-DUNICODE -D_UNICODE)

add_subdirectory(tests)
