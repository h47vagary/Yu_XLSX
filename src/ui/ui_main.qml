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
        color: "#f0f0f0" // 简洁背景
        border.color: "#cccccc"
        radius: 6

        ColumnLayout {
            anchors.centerIn: parent
            spacing: 20
            width: parent.width * 0.8

            RowLayout {
                spacing: 10
                TextField {
                    id: sourceFile
                    placeholderText: "选择数据源文件"
                    Layout.fillWidth: true
                }
                Button {
                    text: "浏览"
                    onClicked: fileDialogSource.open()
                }
            }

            RowLayout {
                spacing: 10
                TextField {
                    id: targetFile
                    placeholderText: "选择目标文件"
                    Layout.fillWidth: true
                }
                Button {
                    text: "浏览"
                    onClicked: fileDialogTarget.open()
                }
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

        GroupBox {
            title: "单价配置"
            Layout.fillWidth: true

            ColumnLayout {
                spacing: 10

                // 数据模型：数量区间 + 单价
                ListModel {
                    id: priceModel
                }

                // 显示所有规则
                ColumnLayout {
                    Repeater {
                        model: priceModel
                        delegate: RowLayout {
                            spacing: 10

                            TextField {
                                placeholderText: "数量 起"
                                text: model.start
                                Layout.preferredWidth: 80
                                onTextChanged: model.start = text
                            }

                            TextField {
                                placeholderText: "数量 止"
                                text: model.end
                                Layout.preferredWidth: 80
                                onTextChanged: model.end = text
                            }

                            TextField {
                                placeholderText: "单价"
                                text: model.price
                                Layout.preferredWidth: 80
                                onTextChanged: model.price = text
                            }

                            Button {
                                text: "删除"
                                onClicked: priceModel.remove(index)
                            }
                        }
                    }
                }

                // 添加新规则按钮
                Button {
                    text: "添加单价规则"
                    Layout.fillWidth: true
                    onClicked: {
                        priceModel.append({
                            "start": "",
                            "end": "",
                            "price": ""
                        })
                    }
                }
            }
        }


        // 文件选择对话框
        FileDialog {
            id: fileDialogSource
            title: "选择数据源文件"
            nameFilters: ["Excel 文件 (*.xlsx *.xls)"]
            onAccepted: sourceFile.text = fileUrl
        }

        FileDialog {
            id: fileDialogTarget
            title: "选择目标文件"
            nameFilters: ["Excel 文件 (*.xlsx *.xls)"]
            onAccepted: targetFile.text = fileUrl
        }

        MessageDialog {
            id: messageDialog
            title: "结果"
            text: ""
            standardButtons: StandardButton.Ok
        }
    }
}
