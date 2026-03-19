cmake_minimum_required(VERSION 3.10)
project(PreviewExplorer CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(MSVC)
    add_compile_options(/source-charset:utf-8)
endif()

add_compile_definitions(UNICODE _UNICODE)

add_library(PreviewExplorer SHARED
    src/MediaExplorer.cpp
)

target_include_directories(PreviewExplorer PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
    ${CMAKE_CURRENT_SOURCE_DIR}/src/sdk
)

target_link_libraries(PreviewExplorer PRIVATE
    user32 comctl32 comdlg32 shlwapi shell32
)

set_target_properties(PreviewExplorer PROPERTIES SUFFIX ".aux2")
set_target_properties(PreviewExplorer PROPERTIES VS_CHARACTER_SET "Unicode")

set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
