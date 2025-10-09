file(GLOB COMMON_SRC ${CMAKE_CURRENT_SOURCE_DIR}/*.cpp)
message(STATUS "xlsx handle sources: ${COMMON_SRC}")

add_library(xlsx_handle STATIC ${COMMON_SRC})

target_include_directories(xlsx_handle PUBLIC
        ${CMAKE_CURRENT_SOURCE_DIR}
        ${PROJECT_SOURCE_DIR}/include
        ${PROJECT_SOURCE_DIR}/include/OpenXLSX
        ${PROJECT_SOURCE_DIR}/include/OpenXLSX/external
)

# 三方库
# OpenXLSX
set(OPENXLSX_DIR "${PROJECT_SOURCE_DIR}/lib/OpenXLSX")
set(OPENXLSX_LIB "${OPENXLSX_DIR}/libOpenXLSX.a")
set(OPENXLSX_NOWIDE_LIB "${OPENXLSX_DIR}/libnowide.a")

# 链接库
target_link_libraries(xlsx_handle
    PUBLIC
        ${OPENXLSX_LIB}
        ${OPENXLSX_NOWIDE_LIB}
)