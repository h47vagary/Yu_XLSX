#include <windows.h>
#include <QGuiApplication>
#include <QQmlAPplicationEngine>
#include <QQmlContext>
#include <QtQuickControls2/QQuickStyle>
#include "excel_finder_controller.h"

int main(int argc, char *argv[])
{
    SetConsoleOutputCP(CP_UTF8);
    QGuiApplication app(argc, argv);

    // 设置非原生风格
    QQuickStyle::setStyle("Fusion");

    qmlRegisterType<ExcelFinderController>(
        "App.Excel", 1, 0, "ExcelFinderController");

    QQmlApplicationEngine engine;
    engine.loadFromModule("YuXlsx", "Main");


    return app.exec();
}