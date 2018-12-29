#include <QApplication>
#include "mainwindow.h"
//#include "vld.h"

int main(int argc, char *argv[])
{
    qRegisterMetaType<SinglePersonInfoGrp>();
    QApplication a(argc, argv);

    MainWindow mWindow;
    mWindow.show();

    return a.exec();
}
