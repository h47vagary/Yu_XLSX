import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Dialogs

Rectangle {
    id: homepage
    
    width: 660
    height: parent.height

    signal source_path_changed(string path)
    signal finder_path_changed(string path)
    signal output_dir_changed(string path)
    signal execute()

    //color: "#28d878ff"
    Rectangle
    {
        id: file_set
        width: parent.width
        height: parent.height / 2

        FileDialog {
            id: filedialog
            title: "选择文件"
            fileMode: FileDialog.OpenFile   // 单文件
            property var targetField

            onAccepted: {
                console.log("选中的文件:", filedialog.selectedFile)
                targetField.text = filedialog.selectedFile.toString()
            }
        }

        FolderDialog {
            id: dirdialog
            title: "选择文件路径"
            property var targetdir
            
            onAccepted: {
                console.log("选中的目录:", selectedFolder)
                targetdir.text = selectedFolder.toString()
            }
        }

        GridLayout {
            id: grid
            anchors.top: parent.top
            anchors.bottom: parent.bottom
            anchors.horizontalCenter: parent.horizontalCenter
            columns: 3
            Text {
                Layout.preferredWidth: 100
                Layout.preferredHeight: 35
                text: "数据源文件"
                font.pixelSize: 15
                font.family: "Regular"
                horizontalAlignment: Text.AlignHCenter
                verticalAlignment: Text.AlignVCenter
            }
            TextField {
                id: source_text
                Layout.preferredWidth: 250
                Layout.preferredHeight: 35
                placeholderText: "请选择数据源路径文件"
                text: ""
                onTextChanged: {
                    homepage.source_path_changed(source_text.text)
                }
            }
            Button {
                id: source_button
                Layout.preferredWidth: 35
                Layout.preferredHeight:35
                onClicked: {
                    console.log("source_button 按钮被点击")
                    filedialog.targetField  = source_text
                    filedialog.open()
                }
            }

            Text {
                Layout.preferredWidth: 75
                Layout.preferredHeight: 35
                text: "待搜索文件"
                font.pixelSize: 15
                font.family: "Regular"
                horizontalAlignment: Text.AlignHCenter
                verticalAlignment: Text.AlignVCenter
            }
            TextField {
                id: finder_text
                Layout.preferredWidth: 250
                Layout.preferredHeight: 35
                placeholderText: "请选择待搜索路径文件"
                text: ""
                onTextChanged: {
                    homepage.finder_path_changed(finder_text.text)
                }
            }
            Button {
                id: finder_button
                Layout.preferredWidth: 35
                Layout.preferredHeight:35
                onClicked: {
                    console.log("finder_button 按钮被点击")
                    filedialog.targetField = finder_text
                    filedialog.open()
                }
            }

            Text {
                Layout.preferredWidth: 100
                Layout.preferredHeight: 35
                text: "导出文件路径"
                font.pixelSize: 15
                font.family: "Regular"
                horizontalAlignment: Text.AlignHCenter
                verticalAlignment: Text.AlignVCenter
            }
            TextField {
                id: output_dir_text
                Layout.preferredWidth: 250
                Layout.preferredHeight: 35
                placeholderText: "请选择导出文件路径"
                text: ""
                onTextChanged: {
                    homepage.output_dir_changed(output_dir_text.text)
                }
            }
            Button {
                id: output_dir_button
                Layout.preferredWidth: 35
                Layout.preferredHeight:35
                onClicked: {
                    console.log("output_dir_button 按钮被点击")
                    dirdialog.targetdir  = output_dir_text
                    dirdialog.open()
                }
            }
        }
    }

    Rectangle {
        id: output_rect
        width: parent.width
        height: parent.height / 2
        anchors.top: file_set.bottom

        ColumnLayout {
            anchors.fill: parent
            spacing: 8

            // ===== 顶部：控制区 =====
            RowLayout {
                Layout.fillWidth: true
                Layout.preferredHeight: 40
                spacing: 12

                Button {
                    id: start_convert
                    text: "开始转换"
                    Layout.preferredWidth: 100
                    Layout.preferredHeight: 30

                    onClicked: {
                        logArea.append("开始转换...")
                        execute()
                    }
                }

                Item {
                    Layout.fillWidth: true   // 占位，让按钮靠左
                }
            }

            // ===== 下方：日志区（占大部分）=====
            TextArea {
                id: logArea
                Layout.fillWidth: true
                Layout.fillHeight: true
                readOnly: true
                wrapMode: Text.Wrap
                placeholderText: "转换打印信息区域"

                background: Rectangle {
                    color: "#f5f5f5"
                    radius: 6
                    border.color: "#cccccc"
                }
            }
        }
    }


    
}