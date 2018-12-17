#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    controllerp = new Controller();
    connect(this, SIGNAL(Sig_Call_Output(QString&,QDate&)), controllerp, SLOT(Slot_StartProgress(QString&,QDate&)));
    connect(controllerp, SIGNAL(Sig_Controller_Status(int,QString)),
            this, SLOT(Slot_Receive_Info(int,QString)));

}

MainWindow::~MainWindow()
{
    delete ui;

    if (controllerp)
    {
        delete controllerp;
        controllerp = Q_NULLPTR;
    }
}

void MainWindow::startProgress(QString &filePath, QDate date)
{
    emit Sig_Call_Output(filePath, date);
}

//处理线程返回信息
void MainWindow::Slot_Receive_Info(int, QString)
{

}

//读入文件
void MainWindow::on_pB_ReadInXlsx_clicked()
{
    QString fileNameStr = QFileDialog::getOpenFileName(Q_NULLPTR,
                                                       QString::fromLocal8Bit("Open the target"),
                                                       QString::fromLocal8Bit("The file is "),
                                                       QString("*.xlsx"));

    //1217 Here

}
