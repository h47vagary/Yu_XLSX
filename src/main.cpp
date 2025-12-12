#include <QGuiApplication>
#include <QQmlApplicationEngine>
#include <QQmlContext>
#include "excel_finder.h"

// 声明 Qt 生成的资源初始化函数
extern int qInitResources_qml();

class FinderBackend : public QObject
{
    Q_OBJECT
public:
    explicit FinderBackend(QObject *parent = nullptr) : QObject(parent) {}

    Q_INVOKABLE bool runFinder(const QString &source, const QString &target, const QString &output)
    {
        ExcelFinder finder(source.toStdString(), target.toStdString());
        if (!finder.init())
            return false;

        finder.set_tags("消费时间", "车牌号码", "数量");
        finder.set_source_read_range(2, 1, 9, 7);

        if (!finder.execute())
            return false;

        return finder.export_results(output.toStdString());
    }
};

int main(int argc, char *argv[])
{
    QGuiApplication app(argc, argv);

    // 手动初始化资源
    qInitResources_qml();

    FinderBackend backend;

    QQmlApplicationEngine engine;
    engine.rootContext()->setContextProperty("finderBackend", &backend);
    engine.load(QUrl(QStringLiteral("qrc:/ui_main.qml")));

    if (engine.rootObjects().isEmpty())
        return -1;

    return app.exec();
}

#include "main.moc"