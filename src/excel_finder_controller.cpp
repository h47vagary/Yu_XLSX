#include "excel_finder_controller.h"

ExcelFinderController::ExcelFinderController(QObject* parent)
    : QObject(parent)
{
}

void ExcelFinderController::set_source_path(const QString &source_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }
    finder_->set_source_path(source_path.toStdString());
}

void ExcelFinderController::set_target_path(const QString &target_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }
    finder_->set_target_path(target_path.toStdString());
}

Q_INVOKABLE void ExcelFinderController::set_output_path(const QString &output_path)
{
    if (!finder_) {
        finder_ = std::make_unique<ExcelFinder>();
    }
    finder_->set_output_path(output_path.toStdString());
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
