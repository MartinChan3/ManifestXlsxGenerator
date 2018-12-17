#include "controller.h"
#include <QDebug>
Controller::Controller()
{

}

//开始整个流程
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
