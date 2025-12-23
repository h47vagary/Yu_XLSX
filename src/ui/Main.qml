import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Window {
    visible: true
    width: 1800
    height: 960
    title: "XFinder"
    
    SideBar {
        id: sidebar

        onHomebuttonclicked: {
            console.log("Mani.qml on homebutton clicked")
        }

        onSettingbuttonclicked: {
            console.log("Mani.qml on settingbutton clicked")
        }
    }
}