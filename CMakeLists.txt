cmake_minimum_required(VERSION 3.10)
project(MediaExplorerPlus CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(AviUtl2SdkRoot "${CMAKE_CURRENT_SOURCE_DIR}/external/aviutl2-plugin-sdk")
set(AviUtl2SdkIncludeDir "${AviUtl2SdkRoot}/include")
if(NOT EXISTS "${AviUtl2SdkIncludeDir}/aviutl2_sdk/plugin2.h")
    message(FATAL_ERROR
        "AviUtl2 SDK mirror not found. Expected external/aviutl2-plugin-sdk/include/aviutl2_sdk/plugin2.h.")
endif()

if(MSVC)
    add_compile_options(/source-charset:utf-8 /wd4828)
endif()

add_compile_definitions(UNICODE _UNICODE)

add_library(MediaExplorerPlus SHARED
    src/MediaExplorer.cpp
    src/RightPane.cpp
    src/TabsSubclass.cpp
    src/SettingsDialog.cpp
)

target_include_directories(MediaExplorerPlus PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
    ${AviUtl2SdkIncludeDir}
)

target_link_libraries(MediaExplorerPlus PRIVATE
    user32 comctl32 comdlg32 shlwapi shell32 ole32 oleaut32 dwmapi
)

set_target_properties(MediaExplorerPlus PROPERTIES SUFFIX ".aux2")
set_target_properties(MediaExplorerPlus PROPERTIES VS_CHARACTER_SET "Unicode")

set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR})
