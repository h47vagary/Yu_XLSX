#pragma once
#include <QObject>
#include <QString>
#include <QtQml/qqml.h>
#include <memory>

#include "excel_finder.h"
class ExcelFinderController : public QObject
{
    Q_OBJECT

    // 给 QML 绑定的状态
    Q_PROPERTY(
            bool busy
            READ busy
            NOTIFY busyChanged
    )

public:
    explicit ExcelFinderController(QObject *parent = nullptr);

    // 给 QML 调用的方法
    Q_INVOKABLE void set_source_path(const QString &source_path);
    Q_INVOKABLE void set_target_path(const QString &target_path);
    Q_INVOKABLE void set_output_path(const QString &output_path);

    Q_INVOKABLE bool setTags(const QString &data,
                             const QString &car,
                             const QString &num);
                             
    Q_INVOKABLE bool execute();

    Q_INVOKABLE bool exportResults();

    bool busy() const { return busy_; }

signals:
    void busyChanged();

    void finished(bool ok, QString error_msg = "");

private:
    std::unique_ptr<ExcelFinder> finder_;
    bool busy_ = false;

};