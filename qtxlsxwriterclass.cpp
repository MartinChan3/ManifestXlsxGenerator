#include "qtxlsxwriterclass.h"
#include <QStandardPaths>
#include <QDir>

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
    QDateTime currentDateTime = QDateTime::currentDateTime();//时间只可以初始化一次
    SinglePersonInfoGrp mSinglePersonInfoGrp = pack.toSinglePersonInfoGrp();
    emit Sig_CallUiShow(mSinglePersonInfoGrp);

    for (int i = 0; i < mSinglePersonInfoGrp.size(); i++)
    {
        //对每个人来说，每个人的条目按照8行一页的标准进行读取
        int singleManThingsNum = mSinglePersonInfoGrp.at(i).infoPtrs.size();
        int size = (int)( singleManThingsNum / SINGLE_PAPER_LINE);
//        if ( singleManThingsNum > size * SINGLE_PAPER_LINE )
//            size++; //若不足规定条数则视为进1
        if (singleManThingsNum % SINGLE_PAPER_LINE)
        {
            size += 1;
        }

        if (size == 0)
        {
            qDebug() << "Wrong of size and return";
            return false;
        }

        for (int j = 0; j < size; j++)
        {
            QXlsx::Document tSDoc(pathOfTemplateXlsxFile);

            writeSingleDocumentDateTime(tSDoc, pack.dateTime);
            writeSingleDocumentThing(tSDoc,
                                     mSinglePersonInfoGrp.at(i),
                                     i * SINGLE_PAPER_LINE + j);

            ///保存,这里需要修改为正确的名称
            QString folderPath;
            if (!generateDesktopDir( currentDateTime, folderPath))
            {
                emit Sig_Status_Info(SIGNAL_WRONG_FOLDER_GENERATOR, "Wrong generator");
                return false;
            }
            QString tSDocNameStr = folderPath +"/" +
                                   //QString::fromLocal8Bit(mSinglePersonInfoGrp.at(i).nameOfPerson.toLatin1())
                                   mSinglePersonInfoGrp.at(i).nameOfPerson
                                   + "_" + QString::number(j)
                                   + ".xlsx";
            ///1224 debug here
            if (!tSDoc.saveAs(tSDocNameStr))
            {
                emit Sig_Status_Info(SIGNAL_WRONG_CANT_SAVE_XLSX, "Can't save xlsx");
                return false;
            }
        }
    }

    return true;
}

//写入单个文件条目类所有文件
bool QtXlsxWriterClass::writeSingleDocumentThing(QXlsx::Document &sDoc,
                                                 const SinglePersonInfo &spI,
                                                 int startIndex)
{
    int writeSize = (spI.infoPtrs.size() - startIndex) >= SINGLE_PAPER_LINE ?
                    SINGLE_PAPER_LINE :
                    (spI.infoPtrs.size() - startIndex);
    qDebug() << QString("info type num %1 and startIndex is %2").arg(spI.infoPtrs.size()).arg(startIndex);
    qDebug() << "WriteSize" << writeSize;

    const SupplyConstInfoPtrs *tInfosPtr = &(spI.infoPtrs);
    for (int i = 0; i < writeSize; i++)
    {
        SupplyInfoPtr lineInfoPtr = tInfosPtr->at(i + startIndex);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 0, lineInfoPtr->nameOfThing);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 1, lineInfoPtr->nameOfUnit);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 2, lineInfoPtr->number);
        sDoc.write(startRowOfThingData + i, startColOfThingData + 3, lineInfoPtr->unitPrice);
    }

    qDebug() << "WriteSize" << writeSize;
    //剩余的编写为空行
    for (int i = writeSize; i < SINGLE_PAPER_LINE; i++)
    {
        sDoc.write(startRowOfThingData + i, startColOfThingData + 0, QVariant());
        sDoc.write(startRowOfThingData + i, startColOfThingData + 1, QVariant());
        sDoc.write(startRowOfThingData + i, startColOfThingData + 2, QVariant());
        sDoc.write(startRowOfThingData + i, startColOfThingData + 3, QVariant());
    }

    return true;
}

bool QtXlsxWriterClass::writeSingleDocumentDateTime(QXlsx::Document &sDoc,
                                                    const QDate &time)
{
    ///1216 here
    sDoc.write(START_ROW_OF_DATE, START_COL_OF_DATE, time.year());
    sDoc.write(START_ROW_OF_DATE, START_COL_OF_DATE, time.month());
    sDoc.write(START_ROW_OF_DATE, START_COL_OF_DATE, time.day());

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
        QString tNameOfPerson = sDoc.read(tStartRow + i, 1).toString().trimmed();
        QString tNameOfThing  = sDoc.read(tStartRow + i, 2).toString().trimmed();
        QString tNameOfUnit   = sDoc.read(tStartRow + i, 3).toString().trimmed();
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
        else
        {
            Q_ASSERT(true);
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

//在桌面生成指定的文件夹
bool generateDesktopDir(const QDateTime &dateTime,
                        QString& dirPath)
{
    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
    QDir dir(desktopPath + QString::fromLocal8Bit("/生成结果"));
    //检查是否需要顶级目录
    if (!dir.exists())
    {
       if (!dir.mkdir(dir.path()))
       {
           return false;
       }
    }
    QString timeStr = dateTime.toString("yyyyMMdd_hhmmss");
    dir.setPath(dir.path() + "/" + timeStr);
    if (!dir.exists())
    {
        if(!dir.mkdir(dir.path()))
        {
            return false;
        }
    }

    dirPath = dir.path();
    return true;
}

