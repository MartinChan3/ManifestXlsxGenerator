#include "controller.h"

Controller::Controller()
{
    connect(&xlsxWriter, SIGNAL(Sig_Status_Info(int,QString)),
            this, SIGNAL(Sig_Controller_Status(int,QString))); 
}

Controller::~Controller()
{

}

//开始整个流程(读取到输出)
void Controller::Slot_StartProgress(QString FilePathStr, QDate date)
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

    ///1219 Debug here
    //Read-in the date
    QXlsx::Document sDoc(FilePathStr);
    mInfoPack.dateTime = date;
    mInfoPack.mSupplyInfos.clear();
    bool ok = xlsxWriter.readSpecifiedFormatData(sDoc, mInfoPack);
    if (!ok)
    {
        emit Sig_Controller_Status(SIGNAL_WRONG_LOADING_INFO, "Wrong info input");
        return;
    }

    //Try output
    ok = xlsxWriter.startWriteNewExcel(mInfoPack);
    if (!ok)
    {
        emit Sig_Controller_Status(SIGNAL_WRONG_OUPUTING_XLSX, "Wrong when write to the text");
        return;
    }

    emit Sig_Controller_Status(SIGNAL_OUTPUT_OK, QString());
    return;
}
