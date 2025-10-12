file(GLOB COMMON_SRC ${CMAKE_CURRENT_SOURCE_DIR}/*.cpp)
message(STATUS "excel finder sources: ${COMMON_SRC}")

if(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
    add_compile_options(-Wno-deprecated-declarations)
endif()


add_library(excel_finder STATIC ${COMMON_SRC})

target_include_directories(excel_finder PUBLIC
        ${CMAKE_CURRENT_SOURCE_DIR}
        ${PROJECT_SOURCE_DIR}/include
        ${PROJECT_SOURCE_DIR}/src/xlsx_handle
        ${PROJECT_SOURCE_DIR}/src/utils
        ${PROJECT_SOURCE_DIR}/src/configfile
)

target_link_libraries( excel_finder
    PRIVATE xlsx_handle
    PRIVATE configfile
    PRIVATE utils
)