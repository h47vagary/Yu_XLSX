import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Dialogs

Rectangle {
    id: homepage
    
    width: 1510
    height: parent.height

    color: "#28d878ff"

    FileDialog {
        id: filedialog
        title: "选择文件"
        property var targetField
        nameFilters: [
            "All files (*)"
        ]
        fileMode: FileDialog.OpenFile   // 单文件
        onAccepted: {
            console.log("选中的文件:", filedialog.selectedFile)
            targetField.text = filedialog.selectedFile.toString()
        }
    }

    GridLayout {
        id: grid

        anchors.top: parent.top
        anchors.bottom: parent.bottom
        anchors.horizontalCenter: parent.horizontalCenter

        columns: 2
        TextField {
            id: source_text
            Layout.preferredWidth: 500
            Layout.preferredHeight: 50
            placeholderText: "请选择数据源路径文件"
            text: ""
        }
        Button {
            id: source_button
            Layout.preferredWidth: 50
            Layout.preferredHeight:50
            onClicked: {
                console.log("source_button 按钮被点击")
                filedialog.targetField  = source_text
                filedialog.open()
            }
        }

        TextField {
            id: finder_text
            Layout.preferredWidth: 500
            Layout.preferredHeight: 50
            placeholderText: "请选择待搜索路径文件"
            text: ""
        }
        Button {
            id: finder_button
            Layout.preferredWidth: 50
            Layout.preferredHeight:50
            onClicked: {
                console.log("finder_button 按钮被点击")
                filedialog.targetField = finder_text
                filedialog.open()
            }
        }
    }
}