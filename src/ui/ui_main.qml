import QtQuick 2.12
import QtQuick.Controls 2.12
import QtQuick.Layouts 1.12
import QtQuick.Dialogs 1.2

ApplicationWindow {
    visible: true
    width: 640
    height: 400
    title: qsTr("Excel 匹配工具")

    Rectangle {
        anchors.fill: parent
        gradient: Gradient {
            GradientStop { position: 0.0; color: "#3a6186" }
            GradientStop { position: 1.0; color: "#89253e" }
        }

        ColumnLayout {
            anchors.centerIn: parent
            spacing: 20
            width: parent.width * 0.8

            TextField {
                id: sourceFile
                placeholderText: "数据源文件路径"
                Layout.fillWidth: true
            }

            TextField {
                id: targetFile
                placeholderText: "目标文件路径"
                Layout.fillWidth: true
            }

            TextField {
                id: outputFile
                placeholderText: "输出结果文件路径"
                Layout.fillWidth: true
                text: "../doc/find_result.xlsx"
            }

            Button {
                text: "开始转换"
                Layout.fillWidth: true
                onClicked: {
                    var ok = finderBackend.runFinder(
                        sourceFile.text,
                        targetFile.text,
                        outputFile.text
                    )
                    if (ok)
                        messageDialog.text = "转换完成！"
                    else
                        messageDialog.text = "转换失败，请检查文件路径。"
                    messageDialog.open()
                }
            }
        }

        MessageDialog {
            id: messageDialog
            title: "结果"
            text: ""
            standardButtons: StandardButton.Ok
        }
    }
}
