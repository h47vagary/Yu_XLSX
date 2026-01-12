add_library(configfile STATIC
                        config_file.cpp
                        filepath_config.cpp
                        output_config.cpp)

target_include_directories(configfile PUBLIC
        ${CMAKE_CURRENT_SOURCE_DIR}
        ${PROJECT_SOURCE_DIR}/include
)

# 链接 
target_link_libraries(configfile
    utils
)