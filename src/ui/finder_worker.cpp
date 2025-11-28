// finder_worker.cpp
#include "finder_worker.h"

// A2
#define EXCEL_FIND_START_ROW 2 
#define EXCEL_FIND_START_COL 1
// H10
#define EXCEL_FIND_END_ROW 10
#define EXCEL_FIND_END_COL 8

FinderWorker::FinderWorker(const QString &searchFile,
                           const QString &targetFile,
                           const QString &outputFile,
                           QObject *parent)
    : QObject(parent),
      m_searchFile(searchFile),
      m_targetFile(targetFile),
      m_outputFile(outputFile)
{}

void FinderWorker::process()
{
    ExcelFinder finder(m_searchFile.toStdString(), m_targetFile.toStdString());

    if (!finder.init()) {
        emit logMessage("业务初始化失败！");
        emit finished();
        return;
    }

    finder.set_tags("消费时间", "车牌号码", "数量");
    finder.set_source_read_range(EXCEL_FIND_START_ROW, EXCEL_FIND_START_COL, EXCEL_FIND_END_ROW, EXCEL_FIND_END_COL);

    if (!finder.execute()) {
        emit logMessage("查找任务执行失败！");
        emit finished();
        return;
    }

    if (!finder.export_results(m_outputFile.toStdString())) {
        emit logMessage("结果导出失败！");
    } else {
        emit logMessage(QString("任务完成，结果已导出至：%1").arg(m_outputFile));
    }

    emit finished();
}
