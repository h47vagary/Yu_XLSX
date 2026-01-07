#pragma once
#include <QObject>
#include <QString>
#include <QtQml/qqml.h> 
#include <memory>

#include "excel_finder.h"

class ExcelFinderController : public QObject
{
    Q_OBJECT
    QML_ELEMENT          // 核心
    QML_NAMED_ELEMENT(ExcelFinder) // QML 中的名字

    // 给 QML 绑定的状态
    Q_PROPERTY(
            bool busy
            READ busy
            NOTIFY busyChanged
    )

    Q_PROPERTY(
            QString lastError
            READ lastError
            NOTIFY lastErrorChanged
    )

public:
    explicit ExcelFinderController(QObject *parent = nullptr);

    // 给 QML 调用的方法
    Q_INVOKABLE bool init(const QString &source,
                          const QString &target);

    Q_INVOKABLE bool setTags(const QString &data,
                             const QString &car,
                             const QString &num);
                             
    Q_INVOKABLE bool execute();

    Q_INVOKABLE bool exportResults(const QString &outFile);

    bool busy() const { return busy_; }
    QString lastError() const { return last_error_; }

signals:
    void busyChanged();
    void lastErrorChanged();
    void finished(bool ok);

private:
    std::unique_ptr<ExcelFinder> finder_;
    bool busy_ = false;
    QString last_error_;
};