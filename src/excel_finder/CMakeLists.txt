if(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
    add_compile_options(-Wno-deprecated-declarations)
endif()


add_library(excel_finder STATIC 
                        excel_finder.cpp)

target_include_directories(excel_finder PUBLIC
        ${CMAKE_CURRENT_SOURCE_DIR}
        ${PROJECT_SOURCE_DIR}/include
        ${PROJECT_SOURCE_DIR}/src/xlsx_handle
        ${PROJECT_SOURCE_DIR}/src/utils
        ${PROJECT_SOURCE_DIR}/src/configfile
)

target_link_libraries( excel_finder
        xlsx_handle
        configfile
        utils
)