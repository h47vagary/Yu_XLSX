# src/ui/CMakeLists.txt

qt5_add_resources(UI_RESOURCES_RCC qml.qrc)

# 用静态库封装资源
add_library(ui_resources STATIC ${UI_RESOURCES_RCC})

# 导出给上层
set(UI_RESOURCES_LIB ui_resources PARENT_SCOPE)
