import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Rectangle {
    id: sidbar
    visible: true
    width: 290
    height:  parent.height
    color: "#ffffff"
    
    signal homebuttonclicked()
    signal settingbuttonclicked()

    property int selectedIndex: -1  // 当前选中按钮， -1 表示没有

    Rectangle {
        id: horizontal_line
        width: parent.width
        height: 1
        color: "#1f1f1f41"
        anchors.top: icon_title.bottom
        anchors.left: parent.left
    }

    Rectangle {
        id: icon_title
        width: parent.width
        height: 90
        color: "#ffffff"
        Row {
            anchors.centerIn: parent
            spacing: 12

            Image {
                source: "qrc:/qt/qml/YuXlsx/src/ui/resource/icon/rabbit.png"
                width: 32
                height: 32
                fillMode: Image.PreserveAspectFit
            }

            Text {
                text: "XFinder"
                font.family: "Helvetica"
                font.pixelSize: 30
                font.bold: true
                verticalAlignment: Text.AlignVCenter
            }
        }
    }

    Rectangle {
        id: vertical_line
        height: parent.height
        width: 1
        color: "#1f1f1f41"
        anchors.top: parent.top
        anchors.left: parent.right
    }

    ColumnLayout {
        anchors.top: icon_title.bottom
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.topMargin: 16
        spacing: 12

        Button {
            id: homepage
            Layout.preferredWidth: 260
            Layout.preferredHeight: 60
            Layout.alignment: Qt.AlignCenter
            spacing: 12
            text: "首页"
            font.pixelSize: 20
            font.family: "Regular"
            font.bold: sidbar.selectedIndex === 0 ? true : false
            icon.source: "qrc:/qt/qml/YuXlsx/src/ui/resource/icon/home.png"
            background: Rectangle {
                color: sidbar.selectedIndex === 0 ? "#444745ff":"#ffffff"
                border.width: 1
                border.pixelAligned: true
                radius: 12
            }
            onClicked: {
                sidbar.selectedIndex = 0
                console.log("首页按钮被点击")
                sidbar.homebuttonclicked()
            }
        }
        Button {
            id: setting
            Layout.preferredWidth: 260
            Layout.preferredHeight: 60
            Layout.alignment: Qt.AlignCenter
            spacing: 12
            text: "设置"
            font.pixelSize: 20
            font.family: "Regular"
            font.bold: sidbar.selectedIndex === 1 ? true : false
            icon.source: "qrc:/qt/qml/YuXlsx/src/ui/resource/icon/setting.png"
            background: Rectangle {
                color: sidbar.selectedIndex === 1 ? "#4a4d4cff":"#ffffff"
                border.width: 1
                border.pixelAligned: true
                radius: 12
            }
            onClicked: {
                sidbar.selectedIndex = 1
                console.log("设置页按钮被点击")
                sidbar.settingbuttonclicked()
            }
        }
    }
}