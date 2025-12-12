cmake_minimum_required(VERSION 3.10)
project(yu_xlsx)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_ON ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_PREFIX_PATH "D:/QT/5.12.9/mingw73_64/lib/cmake")
find_package(Qt5 5.12 REQUIRED COMPONENTS 
    Core
    Gui
    Qml
    Quick 
    QuickControls2
)

include_directories(
    ${PROJECT_SOURCE_DIR}/include
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX/external
    ${PROJECT_SOURCE_DIR}/src/excel_finder
    ${PROJECT_SOURCE_DIR}/src/configfile
    ${PROJECT_SOURCE_DIR}/src/xlsx_handle
    ${PROJECT_SOURCE_DIR}/src/utils
)

add_subdirectory(src/ui)
add_subdirectory(src/excel_finder)
add_subdirectory(src/xlsx_handle)
add_subdirectory(src/configfile)
add_subdirectory(src/utils)

add_executable(yu_xlsx src/main.cpp)

set(OPENXLSX_DIR "${PROJECT_SOURCE_DIR}/lib/OpenXLSX")
set(OPENXLSX_LIB "${OPENXLSX_DIR}/libOpenXLSX.a")
set(OPENXLSX_NOWIDE_LIB "${OPENXLSX_DIR}/libnowide.a")

target_link_libraries(yu_xlsx
    PRIVATE
        ui_lib
        Qt5::Core
        Qt5::Gui
        Qt5::Qml
        Qt5::Quick
        Qt5::QuickControls2
        excel_finder
        xlsx_handle
        configfile
        utils
        ${OPENXLSX_LIB}
        ${OPENXLSX_NOWIDE_LIB}
)