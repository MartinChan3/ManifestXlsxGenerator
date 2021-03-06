#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>

//New Version

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
    connect(controllerp, SIGNAL(Sig_CallUiShow2(SinglePersonInfoGrp)),
            this, SLOT(SlotShowdatainui(SinglePersonInfoGrp)));
    threadControlp->start();

    mTextBroswerp = ui->textBrowser;

    //初始化日期
    QDate today = QDate::currentDate();
    ui->lEofYear->setText(QString::number(today.year()));
    ui->lEofMonth->setText(QString::number(today.month()));
    ui->lEofDay->setText(QString::number(today.day()));

//#define TEST_XLSX
#ifdef  TEST_XLSX
    emit Sig_Call_Output(QString::fromLocal8Bit("./template/test.xlsx"),
                         QDate::currentDate());
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
    case SIGNAL_OUTPUT_OK:
        qDebug() << QString::fromLocal8Bit("整个流程读取无误");
        break;
    default:
        qDebug() << "There is an error happened and please check infos";
        qDebug() << "Info:" << information;
        break;
    }
}

///待写，输出内容
void MainWindow::SlotShowdatainui(SinglePersonInfoGrp grp)
{
    mTextBroswerp->clear();
    for (int i = 0; i < grp.size(); i++)
    {
        QString fstLine;
        fstLine = QString::fromLocal8Bit("姓名：") +
                  grp.at(i).nameOfPerson +
                  QString::fromLocal8Bit("拥有%1条记录").arg(grp.at(i).infoPtrs.size());
        mTextBroswerp->append(fstLine);
        for (int j = 0; j < grp.at(i).infoPtrs.size(); j++)
        {
            const SupplyInfo *cPtr = grp.at(i).infoPtrs.at(j);
            QString lineStr = QString::fromLocal8Bit("    %1 物品：").arg(j) + cPtr->nameOfThing + QString(" ") +
                    QString::fromLocal8Bit("单位：") + cPtr->nameOfUnit + QString(" ") +
                    QString::fromLocal8Bit("数量：") + QString::number(cPtr->number) + QString(" ") +
                    QString::fromLocal8Bit("单价：") + QString::number(cPtr->unitPrice);
            mTextBroswerp->append(lineStr);
        }

        mTextBroswerp->append(QString());
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
