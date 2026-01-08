#include "excel_finder_controller.h"
#include <QUrl>

static std::string urlToLocalPath(const std::string& url)
{
    QUrl qurl(QString::fromStdString(url));
    if (qurl.isLocalFile())
        return qurl.toLocalFile().toStdString();
    return url; // 兜底
}

ExcelFinderController::ExcelFinderController(QObject* parent)
    : QObject(parent)
{
}

void ExcelFinderController::set_source_path(const QString &source_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }
    std::string localPath =
        urlToLocalPath(source_path.toStdString());

    finder_->set_source_path(localPath);
}

void ExcelFinderController::set_target_path(const QString &target_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }

    std::string localPath = 
        urlToLocalPath(target_path.toStdString());
    finder_->set_target_path(localPath);
}

Q_INVOKABLE void ExcelFinderController::set_output_path(const QString &output_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }
    
    std::string localPath = 
        urlToLocalPath(output_path.toStdString());
    finder_->set_output_path(localPath);
}

bool ExcelFinderController::setTags(const QString& data,
                                    const QString& car,
                                    const QString& num)
{
    if (!finder_) return false;

    finder_->set_tags(data.toStdString(),
                      car.toStdString(),
                      num.toStdString());
    return true;
}

bool ExcelFinderController::execute()
{
    if (!finder_) return false;

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
    if (!finder_) return false;
    return finder_->export_results();
}
