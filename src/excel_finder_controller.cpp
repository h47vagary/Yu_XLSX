#include "excel_finder_controller.h"

ExcelFinderController::ExcelFinderController(QObject* parent)
    : QObject(parent)
{
}

bool ExcelFinderController::init(const QString& source,
                                 const QString& target)
{
    finder_ = std::make_unique<ExcelFinder>(
        source.toStdString(),
        target.toStdString());

    if (!finder_->init()) {
        last_error_ = "Excel 文件打开失败";
        emit lastErrorChanged();
        return false;
    }
    return true;
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

bool ExcelFinderController::exportResults(const QString& outFile)
{
    if (!finder_) return false;
    return finder_->export_results(outFile.toStdString());
}
