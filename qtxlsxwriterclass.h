#ifndef QTXLSXWRITERCLASS_H
#define QTXLSXWRITERCLASS_H

#include <QObject>
#include <QVector>
#include "xlsxdocument.h"
#include <QFile>
#include <QDate>

#define SINGLE_PAPER_LINE 8
#define START_ROW 1
#define START_COL 1

enum {
    SIGNAL_OK = 0,    //信息无问题，同时附加信息
    SIGNAL_WRONG = 1
};

typedef struct _SUPPLYINFO{
    QString nameOfPerson;
    QString nameOfThing;
    QString nameOfUnit;
    float number;
    float unitPrice;
}SupplyInfo;
typedef QVector<SupplyInfo> SupplyInfos;

typedef const SupplyInfo* SupplyInfoPtr;
typedef QVector<SupplyInfoPtr> SupplyConstInfoPtrs;

SupplyInfoPtr toPtr(SupplyInfo&);

//单个人信息
typedef struct _SINGLEPERSONINFO{
    QString nameOfPerson;
    SupplyConstInfoPtrs infoPtrs;
}SinglePersonInfo;
typedef QVector<SinglePersonInfo> SinglePersonInfoGrp;

//总信息包
typedef struct _SUPPLYALLPACK{
    QDate dateTime;
    SupplyInfos mSupplyInfos;

    SinglePersonInfoGrp toSinglePersonInfoGrp() const;
}SupplyAllPack;



class QtXlsxWriterClass : public QObject
{
    Q_OBJECT

public:
    QtXlsxWriterClass();

private:
    QString pathOfTemplateXlsxFile;

    int  startRow;
    int  startCol;

public slots:
    bool startWriteNewExcel(const SupplyAllPack&pack);
    bool writeSingleDocumentThing(QXlsx::Document& sDoc,
                                  const SinglePersonInfo& spI,
                                  int startIndex);
    bool writeSingleDocumentDateTime(QXlsx::Document &sDoc,
                                     const QDate &time);



signals:
    bool Sig_Status_Info(int, QString);
};

#endif // QTXLSXWRITERCLASS_H
