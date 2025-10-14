#ifndef MAIN_WINDOW_H
#define MAIN_WINDOW_H

#include <QMainWindow>
#include <QLineEdit>
#include <QPushButton>
#include <QTextEdit>
#include "excel_finder.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void browse_search_file();
    void browse_target_file();
    void browse_output_folder();
    void run_search();

private:
    Ui::MainWindow *ui;

    QString fixed_output_filename() const { return "result.xlsx"; }
};

#endif // MAIN_WINDOW_H
