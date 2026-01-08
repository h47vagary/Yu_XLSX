import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import App.Excel 1.0

Window {
    visible: true
    width: 800
    height: 400
    title: "XFinder"
    
    ExcelFinderController {
        id: finder

        onFinished: (ok, error_msg) => {
            console.log(ok, error_msg)
        }
    }

    // Button {
    //     text: "开始查找"
    //     onClicked: {
    //         finder.init(
    //             "E:\\Code\\Yu_XLSX\\doc\\new\\福邦1.xlsx",
    //             "E:\\Code\\Yu_XLSX\\doc\\new\\福邦14-16交易详情255.xlsx"
    //         )
    //         finder.setTags("消费时间", "车牌号码", "数量")
    //         finder.execute()
    //     }
    // }

    SideBar {
        id: sidebar

        onHomebuttonclicked: {
            console.log("Mani.qml on homebutton clicked")
            homepage.visible = true;
            configpage.visible = false;
        }

        onSettingbuttonclicked: {
            console.log("Mani.qml on settingbutton clicked")
            homepage.visible = false;
            configpage.visible = true;
        }
    }

    HomePage {
        id: homepage
        x: sidebar.width + 1
        visible: false

        onSource_path_changed: (path) => {
            console.log("Main.qml 收到 source_path_changed 信号，路径为:", path)
            finder.set_source_path(path)
        }

        onFinder_path_changed: (path) => {
            console.log("Main.qml 收到 finder_path_changed 信号，路径为:", path)
            finder.set_target_path(path)
        }

        onOutput_dir_changed: (path) => {
            console.log("Main.qml 收到 output_dir_changed 信号，路径为:", path)
            finder.set_output_path(path)
        }

        onExecute: () => {
            console.log("Main.qml 收到 execute 信号，开始执行查找")
            finder.setTags("消费时间", "车牌号码", "数量")
            finder.execute()
            finder.exportResults()
            
        }
    }

    ConfigPage {
        id: configpage
        x: sidebar.width + 1
        visible: false
    }
}