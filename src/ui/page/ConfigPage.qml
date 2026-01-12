import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Rectangle {
    id: configpage
    width: 660
    height: parent.height
    color: "#f6f6f6"

    function reloadPriceRules() {
        priceModel.clear()

        const rules = finder.priceRules()
        console.log("load price rules:", rules.length)

        for (let i = 0; i < rules.length; ++i) {
            priceModel.append({
                min:   Number(rules[i].min),
                max:   Number(rules[i].max),
                price: Number(rules[i].price)
            })
        }
    }


    Component.onCompleted: reloadPriceRules()

    Connections {
        target: finder
        function onPriceRulesChanged() {
            reloadPriceRules()
        }
    }


    ListModel {
        id: priceModel
    }

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: 16
        spacing: 12

        Text {
            text: "数量区间价格配置"
            font.pixelSize: 20
            font.bold: true
        }

        // 表头
        RowLayout {
            spacing: 8
            Text { text: "最小数量"; Layout.preferredWidth: 100 }
            Text { text: "最大数量"; Layout.preferredWidth: 100 }
            Text { text: "价格"; Layout.preferredWidth: 100 }
            Item { Layout.fillWidth: true }
        }

        // 列表
        ListView {
            Layout.fillWidth: true
            Layout.fillHeight: true
            model: priceModel
            spacing: 6
            clip: true

            delegate: RowLayout {
                spacing: 8

                TextField {
                    text: min
                    Layout.preferredWidth: 100
                    inputMethodHints: Qt.ImhDigitsOnly
                    onEditingFinished: {
                        priceModel.setProperty(index, "min", Number(text))
                    }
                }

                TextField {
                    text: max
                    Layout.preferredWidth: 100
                    inputMethodHints: Qt.ImhDigitsOnly
                    onEditingFinished: {
                        priceModel.setProperty(index, "max", Number(text))
                    }
                }

                TextField {
                    text: price
                    Layout.preferredWidth: 100
                    inputMethodHints: Qt.ImhFormattedNumbersOnly
                    onEditingFinished: {
                        priceModel.setProperty(index, "price", Number(text))
                    }
                }

                Button {
                    text: "删除"
                    onClicked: priceModel.remove(index)
                }
            }
        }

        // 底部按钮
        RowLayout {
            spacing: 12

            Button {
                text: "+ 添加区间"
                onClicked: {
                    priceModel.append({ min: 0, max: 0, price: 0.0 })
                }
            }

            Item { Layout.fillWidth: true }

            Button {
                text: "保存配置"
                onClicked: {
                    let list = []

                    for (let i = 0; i < priceModel.count; ++i) {
                        const r = priceModel.get(i)
                        list.push({
                            min: Number(r.min),
                            max: Number(r.max),
                            price: Number(r.price)
                        })
                    }

                    console.log("current price rules:")
                    for (let i = 0; i < list.length; ++i) {
                        console.log(list[i].min, list[i].max, list[i].price)
                    }

                    const ok = finder.setPriceRules(list)
                    console.log("set price rules result:", ok)
                }
            }

        }
    }
}
