#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "controller.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private:
    Ui::MainWindow *ui;

public slots:
    void startProgress(QString & filePath, QDate date = QDate::currentDate());

signals:
    void Sig_Call_Output(QString&, QDate&);

};

#endif // MAINWINDOW_H
