#ifndef QTXLSXWRITERCLASS_H
#define QTXLSXWRITERCLASS_H

#include <QObject>
#include <QVector>
#include "xlsxdocument.h"
#include <QFile>

#define INTERV

enum {
    SIGNAL_OK = 0,
    SIGNAL_WRONG = 1
};

typedef struct _SUPPLYINFO{
    QString name;
    int number;
    float unitPrice;
}SupplyInfo;
typedef QVector<SupplyInfo> SupplyInfos;

typedef struct _SUPPLYALLPACK{
    QDate dateTime;
    SupplyInfos mSupplyInfos;
}SupplyAllPack;



class QtXlsxWriterClass
{
public:
    QtXlsxWriterClass();

private:
    QString pathOfTemplateXlsxFile;

public slots:
    bool startWriteNewExcel(const SupplyAllPack&pack);
    bool writeSingleDocumentThing(QXlsx::Document& sDoc,
                                  const SupplyAllPack& pack,
                                  int startThing,
                                  );

signals:
    bool Sig_Status_Info(int, QString);
};

#endif // QTXLSXWRITERCLASS_H
