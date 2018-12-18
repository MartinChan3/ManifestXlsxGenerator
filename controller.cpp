#include "controller.h"

Controller::Controller()
{
    threadControlp = new QThread();

    this->moveToThread(threadControlp);
    connect(threadControlp, SIGNAL(finished()), this, SLOT(deleteLater()));
    connect(&xlsxWriter, SIGNAL(Sig_Status_Info(int,QString)),
            this, SIGNAL(Sig_Controller_Status(int,QString)));
    threadControlp->start();
}

Controller::~Controller()
{
    threadControlp->quit();
    threadControlp->wait();

    if (threadControlp)
    {
        delete threadControlp;
        threadControlp = Q_NULLPTR;
    }
}

//开始整个流程(读取到输出)
void Controller::Slot_StartProgress(QString &FilePathStr, QDate &date)
{
    if (FilePathStr == inFilePath && date == inDate)
    {
        return;
    }

    QFile tFile(FilePathStr);
    if (!tFile.exists())
    {
        qDebug() << "No valid document was founded";
        return;
    }

    //Read-in the date
    QXlsx::Document sDoc(FilePathStr);
    mInfoPack.dateTime = date;
    mInfoPack.mSupplyInfos.clear();
    bool ok = xlsxWriter.readSpecifiedFormatData(sDoc, mInfoPack);
    if (!ok)
    {
        return;
    }

    //Try output
    ok = xlsxWriter.startWriteNewExcel(mInfoPack);
    if (!ok)
    {
        return;
    }
}
