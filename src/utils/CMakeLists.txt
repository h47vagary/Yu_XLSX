add_library(utils STATIC 
                time_utils.cpp)

target_include_directories(utils PUBLIC 
                            ${PROJECT_SOURCE_DIR}/include
                            ${PROJECT_SOURCE_DIR}/lib                            
)
