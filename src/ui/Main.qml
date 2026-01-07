import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import YuXlsx 1.0


Window {
    visible: true
    width: 800
    height: 400
    title: "XFinder"
    
    ExcelFinder {
        id: finder

        onFinished: (ok) => {
            console.log("查找完成:", ok)
        }
    }

    Button {
        text: "开始查找"
        onClicked: {
            finder.init(
                "E:\\Code\\Yu_XLSX\\doc\\new\\福邦1.xlsx",
                "E:\\Code\\Yu_XLSX\\doc\\new\\福邦14-16交易详情255.xlsx"
            )
            finder.setTags("消费时间", "车牌号码", "数量")
            finder.execute()
        }
    }

    // SideBar {
    //     id: sidebar

    //     onHomebuttonclicked: {
    //         console.log("Mani.qml on homebutton clicked")
    //         homepage.visible = true;
    //     }

    //     onSettingbuttonclicked: {
    //         console.log("Mani.qml on settingbutton clicked")
    //         homepage.visible = false;
    //     }
    // }

    // HomePage {
    //     id: homepage
    //     x: sidebar.width + 1
    //     visible: false
    // }
}