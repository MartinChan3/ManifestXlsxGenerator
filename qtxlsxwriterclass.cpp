#include "qtxlsxwriterclass.h"

QtXlsxWriterClass::QtXlsxWriterClass()
{

}

bool QtXlsxWriterClass::startWriteNewExcel(const SupplyAllPack &pack)
{
    int packSize = pack.mSupplyInfos.size();

    if (!packSize)
    {
        emit Sig_Status_Info(SIGNAL_WRONG, QStringLiteral("无有效的条目输入"));
        return false;
    }

    if (!QFile(pathOfTemplateXlsxFile).exists())
    {
        emit Sig_Status_Info(SIGNAL_WRONG, QStringLiteral("无有效的模板文件"));
        return false;
    }

    int fileNums = (int)(packSize / 4);

    //创建xlsx文件
    for (int i = 0; i < fileNums; i++)
    {
        QXlsx::Document sDocument(pathOfTemplateXlsxFile);



    }
}
