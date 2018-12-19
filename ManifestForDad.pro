#-------------------------------------------------
#
# Project created by QtCreator 2018-12-14T15:24:38
#
#-------------------------------------------------
include($$PWD/xlsx/qtxlsx.pri)

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ManifestForDad
TEMPLATE = app

win32:CONFIG(release, debug|release): DESTDIR = ../OUTPUT/release
else:win32:CONFIG(debug, debug|release): DESTDIR = ../OUTPUT/debug

SOURCES += \
    main.cpp \
    mainwindow.cpp \
    qtxlsxwriterclass.cpp \
    controller.cpp

FORMS += \
    mainwindow.ui

HEADERS += \
    mainwindow.h \
    qtxlsxwriterclass.h \
    controller.h

DISTFILES += \
    README.md

#Config vld for debugging
win32 {
    CONFIG(debug, debug|release) {
        DEFINES += VLD_MODULE
        VLD_PATH = "C:\Program Files (x86)\Visual Leak Detector"
        INCLUDEPATH += $${VLD_PATH}/include
        LIBS += -L$${VLD_PATH}/lib/Win32 -lvld
    }
}
