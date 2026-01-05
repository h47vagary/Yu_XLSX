import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Window {
    visible: true
    width: 800
    height: 400
    title: "XFinder"
    
    SideBar {
        id: sidebar

        onHomebuttonclicked: {
            console.log("Mani.qml on homebutton clicked")
            homepage.visible = true;
        }

        onSettingbuttonclicked: {
            console.log("Mani.qml on settingbutton clicked")
            homepage.visible = false;
        }
    }

    HomePage {
        id: homepage
        x: sidebar.width + 1
        visible: false
    }
}