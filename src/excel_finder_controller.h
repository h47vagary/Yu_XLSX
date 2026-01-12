#pragma once
#include <QObject>
#include <QString>
#include <QTimer>
#include <QtConcurrent>
#include <QtQml/qqml.h>
#include <memory>

#include "excel_finder.h"
#include "excel_finder_config_handler.h"
#include "config_manager.h"

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
    ~ExcelFinderController();

    // 给 QML 调用的方法
    Q_INVOKABLE void set_source_path(const QString &source_path);
    Q_INVOKABLE void set_target_path(const QString &target_path);
    Q_INVOKABLE void set_output_path(const QString &output_path);

    Q_INVOKABLE QString sourcePath() const;
    Q_INVOKABLE QString targetPath() const;
    Q_INVOKABLE QString outputPath() const;


    Q_INVOKABLE bool setTags(const QString &data,
                             const QString &car,
                             const QString &num);
                             
    Q_INVOKABLE bool execute();

    Q_INVOKABLE bool exportResults();

    Q_INVOKABLE bool setPriceRules(const QVariantList &price_rules);
    Q_INVOKABLE QVariantList priceRules() const;


    bool busy() const { return busy_; }

    inline void scheduleSave() {
        save_timer_.start(); // 自动重置计时
    }

signals:
    void busyChanged();

    void finished(bool ok, QString error_msg = "");

    void priceRulesChanged();

private:
    std::unique_ptr<ExcelFinder> finder_;
    std::unique_ptr<ExcelFinderConfigHandler> config_handler_;
    bool busy_ = false;
    QTimer save_timer_;
};