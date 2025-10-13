cmake_minimum_required(VERSION 3.10)
project(yu_xlsx)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_ON ON)


#Qt 所需
set(CMAKE_PREFIX_PATH "D:/QT/5.12.9/mingw73_64/lib/cmake")
find_package(Qt5 REQUIRED COMPONENTS Quick QuickControls2)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

# 头文件
include_directories(
    ${PROJECT_SOURCE_DIR}/include
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX/external
    ${PROJECT_SOURCE_DIR}/lib/jsoncpp/include
    ${PROJECT_SOURCE_DIR}/src/excel_finder
    ${PROJECT_SOURCE_DIR}/src/configfile
    ${PROJECT_SOURCE_DIR}/src/xlsx_handle
    ${PROJECT_SOURCE_DIR}/src/utils
)

# 子模块
add_subdirectory(src/ui)
add_subdirectory(src/excel_finder)
add_subdirectory(src/xlsx_handle)
add_subdirectory(src/configfile)
add_subdirectory(src/utils)

# 主程序
add_executable(yu_xlsx src/main.cpp)

# 三方库
# OpenXLSX
set(OPENXLSX_DIR "${PROJECT_SOURCE_DIR}/lib/OpenXLSX")
set(OPENXLSX_LIB "${OPENXLSX_DIR}/libOpenXLSX.a")
set(OPENXLSX_NOWIDE_LIB "${OPENXLSX_DIR}/libnowide.a")

# 链接库
target_link_libraries(yu_xlsx
    PRIVATE
        ${OPENXLSX_LIB}
        ${OPENXLSX_NOWIDE_LIB}
        Qt5::Quick
        Qt5::QuickControls2
        excel_finder
        xlsx_handle
        configfile
        utils
        ${UI_RESOURCES_LIB}    # 链接资源库
)