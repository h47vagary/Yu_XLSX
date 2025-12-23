cmake_minimum_required(VERSION 3.10)
project(yu_xlsx)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_ON ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_PREFIX_PATH "D:/QT6/6.10.1/mingw_64")
find_package(Qt6 REQUIRED COMPONENTS Quick Quick3D)
qt_standard_project_setup(REQUIRES 6.8)

include_directories(
    ${PROJECT_SOURCE_DIR}/include
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX
    ${PROJECT_SOURCE_DIR}/include/OpenXLSX/external
    ${PROJECT_SOURCE_DIR}/src/excel_finder
    ${PROJECT_SOURCE_DIR}/src/configfile
    ${PROJECT_SOURCE_DIR}/src/xlsx_handle
    ${PROJECT_SOURCE_DIR}/src/utils
)

# =========================
# 子模块
# =========================
add_subdirectory(src/excel_finder)
add_subdirectory(src/xlsx_handle)
add_subdirectory(src/configfile)
add_subdirectory(src/utils)

# =========================
# 可执行程序
# =========================
qt_add_executable(yu_xlsx
    src/main.cpp
)

# ===== QML 模块 =====
qt_add_qml_module(yu_xlsx
    URI YuXlsx
    VERSION 1.0
    QML_FILES
        src/ui/Main.qml
        src/ui/component/SideBar.qml
        #ui/page/XXX.qml
    RESOURCES
        src/ui/resource/icon/home.png
        src/ui/resource/icon/setting.png
        src/ui/resource/icon/rabbit.png
)


set(OPENXLSX_DIR "${PROJECT_SOURCE_DIR}/lib/OpenXLSX")
set(OPENXLSX_LIB "${OPENXLSX_DIR}/libOpenXLSX.a")
set(OPENXLSX_NOWIDE_LIB "${OPENXLSX_DIR}/libnowide.a")

target_link_libraries(yu_xlsx
    PRIVATE
        Qt6::Quick
        Qt6::Quick3D
        excel_finder
        xlsx_handle
        configfile
        utils
        ${OPENXLSX_LIB}
        ${OPENXLSX_NOWIDE_LIB}
)