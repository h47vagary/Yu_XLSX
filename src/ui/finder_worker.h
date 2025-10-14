// finder_worker.h
#ifndef FINDER_WORKER_H
#define FINDER_WORKER_H

#include <QObject>
#include <QString>
#include "excel_finder.h"

class FinderWorker : public QObject
{
    Q_OBJECT
public:
    FinderWorker(const QString &searchFile,
                 const QString &targetFile,
                 const QString &outputFile,
                 QObject *parent = nullptr);

signals:
    void logMessage(const QString &msg);
    void finished();

public slots:
    void process();

private:
    QString m_searchFile;
    QString m_targetFile;
    QString m_outputFile;
};

#endif // FINDER_WORKER_H
