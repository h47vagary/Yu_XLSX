#include "main_window.h"
#include "ui_main_window.h"
#include <QFileDialog>
#include <QDir>
#include <QString>
#include <QThread>
#include "finder_worker.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    connect(ui->btn_browse_search, &QPushButton::clicked, this, &MainWindow::browse_search_file);
    connect(ui->btn_browse_target, &QPushButton::clicked, this, &MainWindow::browse_target_file);
    connect(ui->btn_browse_output, &QPushButton::clicked, this, &MainWindow::browse_output_folder);
    connect(ui->btn_run, &QPushButton::clicked, this, &MainWindow::run_search);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::browse_search_file()
{
    QString file = QFileDialog::getOpenFileName(this, "选择待搜索文件");
    if (!file.isEmpty())
        ui->edit_search_file->setText(file);
}

void MainWindow::browse_target_file()
{
    QString file = QFileDialog::getOpenFileName(this, "选择数据源文件");
    if (!file.isEmpty())
        ui->edit_target_file->setText(file);
}

void MainWindow::browse_output_folder()
{
    QString folder = QFileDialog::getExistingDirectory(this, "选择导出文件夹");
    if (!folder.isEmpty())
        ui->edit_output_file->setText(QDir(folder).filePath(fixed_output_filename()));
}

void MainWindow::run_search()
{
    QString search_file = ui->edit_search_file->text();
    QString target_file = ui->edit_target_file->text();
    QString output_file = ui->edit_output_file->text();

    ui->txt_output->clear();

    if (search_file.isEmpty() || target_file.isEmpty() || output_file.isEmpty()) {
        ui->txt_output->append("请填写所有路径！");
        return;
    }

    // 禁用按钮
    ui->btn_run->setEnabled(false);
    ui->txt_output->clear();
    ui->txt_output->append("搜索中...请等待...");

    // 创建线程和 Worker
    QThread *thread = new QThread;
    FinderWorker *worker = new FinderWorker(search_file, target_file, output_file);
    worker->moveToThread(thread);

    // 线程启动时执行 process()
    connect(thread, &QThread::started, worker, &FinderWorker::process);

    // 日志通过信号回到主线程
    connect(worker, &FinderWorker::logMessage, this, [this](const QString &msg){
        ui->txt_output->append(msg);
    }, Qt::QueuedConnection);

    // 完成后清理线程和 Worker
    connect(worker, &FinderWorker::finished, thread, &QThread::quit);
    connect(worker, &FinderWorker::finished, worker, &FinderWorker::deleteLater);
    connect(thread, &QThread::finished, thread, &QThread::deleteLater);

    // 线程结束时重新启用按钮
    connect(thread, &QThread::finished, this, [this](){
        ui->btn_run->setEnabled(true);
    }, Qt::QueuedConnection);

    thread->start();
}
