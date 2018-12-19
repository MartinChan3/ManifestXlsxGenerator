#include "qtxlsxwriterclass.h"

QtXlsxWriterClass::QtXlsxWriterClass()
{
    startRowOfThingData = START_ROW;
    startColOfThingData = START_COL;

    startRowOfDate = START_ROW_OF_DATE;
    startColOfDate = START_COL_OF_DATE;

    pathOfTemplateXlsxFile = QString("./template/RuKuDan.xlsx");
}

//总写入函数
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

    //获取并分类全部信息
    SinglePersonInfoGrp mSinglePersonInfoGrp = pack.toSinglePersonInfoGrp();
    for (int i = 0; i < mSinglePersonInfoGrp.size(); i++)
    {
        //对每个人来说，每个人的条目按照8行一页的标准进行读取
        int singleManThingsNum = mSinglePersonInfoGrp.at(i).infoPtrs.size();
        int size = (int)(singleManThingsNum / SINGLE_PAPER_LINE);
        if (singleManThingsNum > size * SINGLE_PAPER_LINE)
            size++; //若不足规定条数则视为进1

        for (int j = 0; j < size; j++)
        {
            QXlsx::Document tSDoc(pathOfTemplateXlsxFile);

            writeSingleDocumentDateTime(tSDoc, pack.dateTime);
            writeSingleDocumentThing(tSDoc,
                                     mSinglePersonInfoGrp.at(i),
                                     i * SINGLE_PAPER_LINE + j);

            ///保存,这里需要修改为正确的名称
            QString tSDocNameStr = QString::fromLocal8Bit(mSinglePersonInfoGrp.at(i).nameOfPerson.toLatin1())
                    + ".xlsx";
            tSDoc.saveAs(tSDocNameStr);
        }
    }
}

//写入单个文件条目类所有文件
bool QtXlsxWriterClass::writeSingleDocumentThing(QXlsx::Document &sDoc,
                                                 const SinglePersonInfo &spI,
                                                 int startIndex)
{
    int writeSize = (spI.infoPtrs.size() - startIndex) >= SINGLE_PAPER_LINE ?
                    SINGLE_PAPER_LINE :
                    (spI.infoPtrs.size() - startIndex);

    const SupplyConstInfoPtrs *tInfosPtr = &(spI.infoPtrs);
    for (int i = 0; i < writeSize; i++)
    {
        SupplyInfoPtr lineInfoPtr = tInfosPtr->at(i + startIndex);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 0, lineInfoPtr->nameOfThing);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 1, lineInfoPtr->nameOfUnit);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 2, lineInfoPtr->number);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 3, lineInfoPtr->unitPrice);
    }

    return true;
}

bool QtXlsxWriterClass::writeSingleDocumentDateTime(QXlsx::Document &sDoc,
                                                    const QDate &time)
{
    ///1216 here
    sDoc.write(START_ROW,START_COL,time.year());
    sDoc.write(START_ROW,START_COL,time.month());
    sDoc.write(START_ROW,START_COL,time.day());

    return true;
}

//读取特定的xlsx文件的内容
bool QtXlsxWriterClass::readSpecifiedFormatData(QXlsx::Document &sDoc, SupplyAllPack &rPack)
{
    QXlsx::CellRange rangeOfDoc = sDoc.dimension();
    int tStartRow = rangeOfDoc.firstRow();
    int tEndRow = rangeOfDoc.lastRow();

    bool tOk1, tOk2;

    for (int i = tStartRow; i < tEndRow; i++)
    {
        ///1219 此处需要根据实际情况修改对应的位置的内容
        QString tNameOfPerson = sDoc.read(tStartRow + i, 1).toString();
        QString tNameOfThing  = sDoc.read(tStartRow + i, 2).toString();
        QString tNameOfUnit   = sDoc.read(tStartRow + i, 3).toString();
        float   tNumber       = sDoc.read(tStartRow + i, 4).toFloat(&tOk1);
        float   tUnitPrice    = sDoc.read(tStartRow + i, 5).toFloat(&tOk2);

        if (tOk1 && tOk2 && !tNameOfPerson.isEmpty() && !tNameOfThing.isEmpty())
        {
            SupplyInfo tSingleInfo;
            tSingleInfo.nameOfPerson = tNameOfPerson;
            tSingleInfo.nameOfThing  = tNameOfThing;
            tSingleInfo.nameOfUnit   = tNameOfUnit;
            tSingleInfo.number       = tNumber;
            tSingleInfo.unitPrice    = tUnitPrice;

            rPack.mSupplyInfos << tSingleInfo;
        }
    }

    if (rPack.mSupplyInfos.size() == 0)
        return false;
    else
        return true;
}

//转换为以人名分类的情况
SinglePersonInfoGrp _SUPPLYALLPACK::toSinglePersonInfoGrp() const
{
    QVector<QString> names;
    for (int i = 0; i < this->mSupplyInfos.size(); i++)
    {
        if (!names.contains(mSupplyInfos.at(i).nameOfPerson))
        {
            names << mSupplyInfos.at(i).nameOfPerson;
        }
    }

    SinglePersonInfoGrp tSinglePersonInfoGrp;
    for (int i = 0; i < names.size(); i++)
    {
        SinglePersonInfo tSinglePersonInfo;
        tSinglePersonInfo.nameOfPerson = names.at(i);
        for (int j = 0; j < mSupplyInfos.size(); j++)
        {
            if (tSinglePersonInfo.nameOfPerson == mSupplyInfos.at(j).nameOfPerson)
            {
                tSinglePersonInfo.infoPtrs.append(&(mSupplyInfos.at(j)));
            }
        }

        tSinglePersonInfoGrp << tSinglePersonInfo;
    }

    return tSinglePersonInfoGrp;
}

SupplyInfoPtr toPtr(const SupplyInfo &info)
{
    return &info;
}
