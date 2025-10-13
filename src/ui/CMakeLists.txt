# # ui 子目录的 CMakeLists.txt
# qt5_add_resources(UI_RESOURCES_RCC
#     qml.qrc
# )

# # 导出给上层 CMakeLists
# set(UI_RESOURCES_RCC ${UI_RESOURCES_RCC} PARENT_SCOPE)

# src/ui/CMakeLists.txt
qt5_add_resources(UI_RESOURCES_RCC
    qml.qrc
)

# 将生成的源文件列表导出给上层
set(UI_RESOURCES_RCC ${UI_RESOURCES_RCC} PARENT_SCOPE)
