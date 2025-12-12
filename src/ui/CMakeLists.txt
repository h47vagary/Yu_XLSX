qt5_add_resources(UI_QML
    qml.qrc
)

add_library(ui_lib STATIC
    ${UI_QML}
)

target_link_libraries(ui_lib
    Qt5::Qml
    Qt5::Quick
    Qt5::QuickControls2
)
