#include "excel_finder_controller.h"
#include <QUrl>
#include <QVariantMap>
#include <QDebug>
#include <QJSValue>
#include <QJSEngine>


static std::string urlToLocalPath(const std::string& url)
{
    QUrl qurl(QString::fromStdString(url));
    if (qurl.isLocalFile())
        return qurl.toLocalFile().toStdString();
    return url;
}

ExcelFinderController::ExcelFinderController(QObject* parent)
    : QObject(parent)
    , finder_(std::make_unique<ExcelFinder>())
{
    // 创建配置处理器（绑定 finder）
    config_handler_ =
        std::make_unique<ExcelFinderConfigHandler>(*finder_);

    // 注册到配置中心
    auto& cfg = ConfigManager::instance();
    cfg.register_configurable(config_handler_.get());

    // 只在第一次 Controller 创建时 load
    cfg.load_all();

    // ===== 无感异步保存初始化 =====
    save_timer_.setSingleShot(true);
    save_timer_.setInterval(500); // 500ms debounce

    connect(&save_timer_, &QTimer::timeout, this, [] {
        QtConcurrent::run([] {
            ConfigManager::instance().save_all();
        });
    });
}

ExcelFinderController::~ExcelFinderController()
{
    ConfigManager::instance().save_all();
}

void ExcelFinderController::set_source_path(const QString &source_path)
{
    if (!finder_ || !config_handler_) return;

    std::string localPath =
        urlToLocalPath(source_path.toStdString());

    finder_->set_source_path(localPath);
    scheduleSave();
}

void ExcelFinderController::set_target_path(const QString &target_path)
{
    if (!finder_ || !config_handler_) return;

    std::string localPath = 
        urlToLocalPath(target_path.toStdString());
    finder_->set_target_path(localPath);
    scheduleSave();
}

Q_INVOKABLE void ExcelFinderController::set_output_path(const QString &output_path)
{
    if (!finder_ || !config_handler_) return;
    
    std::string localPath = 
        urlToLocalPath(output_path.toStdString());
    finder_->set_output_path(localPath);
    scheduleSave();
}

Q_INVOKABLE QString ExcelFinderController::sourcePath() const
{
    if (!finder_ || !config_handler_) return "";

    return finder_->get_source_path().c_str();
}

Q_INVOKABLE QString ExcelFinderController::targetPath() const
{
    if (!finder_ || !config_handler_) return "";

    return finder_->get_target_path().c_str();
}

Q_INVOKABLE QString ExcelFinderController::outputPath() const
{
    if (!finder_ || !config_handler_) return "";

    return finder_->get_output_path().c_str();
}

bool ExcelFinderController::setTags(const QString& data,
                                    const QString& car,
                                    const QString& num)
{
    if (!finder_ || !config_handler_) return false;

    finder_->set_tags(data.toStdString(),
                      car.toStdString(),
                      num.toStdString());
    return true;
}

bool ExcelFinderController::execute()
{
    if (!finder_ || !config_handler_) return false;

    if (finder_->get_source_path().empty() ||
        finder_->get_target_path().empty() ||
        finder_->get_output_path().empty())
    {
        emit finished(false, "路径未配置完整");
        return false;
    }

    busy_ = true;
    emit busyChanged();

    bool ok = finder_->execute();

    busy_ = false;
    emit busyChanged();

    emit finished(ok);
    return ok;
}

bool ExcelFinderController::exportResults()
{
    if (!finder_ || !config_handler_) return false;
    return finder_->export_results();
}

Q_INVOKABLE bool ExcelFinderController::setPriceRules(const QVariantList &price_rules)
{
    qDebug() << "set price rules count:" << price_rules.size();
    if (!finder_ || !config_handler_) return false;

    finder_->clear_price();

    for (const QVariant& v : price_rules)
    {
        qDebug() << "process price rule:" << v;

        QVariantMap map;

        if (v.canConvert<QVariantMap>())
        {
            map = v.toMap();
        }
        else if (v.canConvert<QJSValue>())
        {
            QJSValue js = v.value<QJSValue>();
            map = js.toVariant().toMap();
        }
        else
        {
            qWarning() << "unknown price rule type:" << v;
            continue;
        }

        double min   = map.value("min").toDouble();
        double max   = map.value("max").toDouble();
        double price = map.value("price").toDouble();

        if (!finder_->add_price(min, max, price))
        {
            qWarning() << "add price rule failed:" << min << max << price;
            return false;
        }
    }
    scheduleSave();
    emit priceRulesChanged();
    qDebug() << "add finshed :" << price_rules.size();
    return true;
}

QVariantList ExcelFinderController::priceRules() const
{
    QVariantList list;
    if (!finder_) return list;

    const auto& rules = finder_->get_price_map();

    for (const auto& [range, price_value] : rules)
    {
        QVariantMap map;
        map["min"]   = range.min_quantity;
        map["max"]   = range.max_quantity;
        map["price"] = price_value;
        list.push_back(map);
    }
    return list;
}
