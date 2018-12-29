#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "controller.h"
#include <QThread>
#include <QTextBrowser>

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
    Controller *controllerp;

    QThread *threadControlp;
    QTextBrowser *mTextBroswerp;

public slots:
    void startProgress(QString filePath, QDate date = QDate::currentDate());
    void Slot_Receive_Info(int, QString);

    void SlotShowdatainui(SinglePersonInfoGrp);
signals:
    void Sig_Call_Output(QString, QDate);

private slots:
    void on_pB_ReadInXlsx_clicked();
};

#endif // MAINWINDOW_H
