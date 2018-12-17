#ifndef CONTROLLER_H
#define CONTROLLER_H

#include "qtxlsxwriterclass.h"
#include <QThread>

class Controller : public QObject
{
    Q_OBJECT

public:
    Controller();
    ~Controller();

private:
    QtXlsxWriterClass xlsxWriter;
    QString inFilePath;
    QDate   inDate;
    SupplyAllPack mInfoPack;

    QThread *threadControlp;

public slots:
    void Slot_StartProgress(QString&, QDate&);

signals:
    void Sig_Controller_Status(int, QString);
};

#endif // CONTROLLER_H
