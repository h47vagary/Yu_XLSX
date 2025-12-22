import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Rectangle {
    id: sidbar
    visible: true
    width: 290
    height:  parent.height
    color: "#ffffff"
    
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
        // border.color: "#1f1f1f"
        // border.width: 1
        Text {
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.verticalCenter: parent.verticalCenter
            text: "XFinder"
            font.family: "Helvetica"
            font.pixelSize: 30
            font.bold: true
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
            text: "首页"
        }
        Button {
            id: setting
            Layout.preferredWidth: 260
            Layout.preferredHeight: 60
            Layout.alignment: Qt.AlignCenter
            text: "设置"
        }
    }
}