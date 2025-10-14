#include <QApplication>
#include <windows.h>
#include "main_window.h"

int main(int argc, char *argv[])
{
    SetConsoleOutputCP(CP_UTF8); // 设置控制台输出为 UTF-8
    QApplication app(argc, argv);
    MainWindow w;
    w.show();
    return app.exec();
}
