#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    controllerp = new Controller();
    threadControlp = new QThread();
    controllerp->moveToThread(threadControlp);
    connect(threadControlp, SIGNAL(finished()), this, SLOT(deleteLater()));
    connect(this, SIGNAL(Sig_Call_Output(QString,QDate)),
            controllerp, SLOT(Slot_StartProgress(QString,QDate)));
    connect(controllerp, SIGNAL(Sig_Controller_Status(int,QString)),
            this, SLOT(Slot_Receive_Info(int,QString)));
    threadControlp->start();

    //初始化日期
    QDate today = QDate::currentDate();
    ui->lEofYear->setText(QString::number(today.year()));
    ui->lEofMonth->setText(QString::number(today.month()));
    ui->lEofDay->setText(QString::number(today.day()));

#define TEST_XLSX
#ifdef  TEST_XLSX
    emit Sig_Call_Output(QString::fromLocal8Bit("./template/工作簿1.xlsx"), QDate::currentDate());
#endif
}

MainWindow::~MainWindow()
{
    delete ui;

    threadControlp->quit();
    threadControlp->wait();

    if (threadControlp)
    {
        delete threadControlp;
        threadControlp = Q_NULLPTR;
    }

    if (controllerp)
    {
        delete controllerp;
        controllerp = Q_NULLPTR;
    }
}

void MainWindow::startProgress(QString filePath, QDate date)
{
    emit Sig_Call_Output(filePath, date);
}

//处理线程返回信息
void MainWindow::Slot_Receive_Info(int statusNumber, QString information)
{
    switch (statusNumber)
    {
    case SIGNAL_OK:
        break;
    default:
        qDebug() << "There is an error happened and please check";
        qDebug() << information;
        break;
    }
}

//读入文件
void MainWindow::on_pB_ReadInXlsx_clicked()
{
    QString fileNameStr = QFileDialog::getOpenFileName(Q_NULLPTR,
                                                       QString::fromLocal8Bit("Open the target"),
                                                       QString("./template"),
                                                       QString("*.xlsx"));

    qDebug() << fileNameStr;
    QDate time(ui->lEofYear->text().toInt(), ui->lEofMonth->text().toInt(), ui->lEofDay->text().toInt());
    emit Sig_Call_Output(fileNameStr, time);//等待次线程自动返回全部内容
}
