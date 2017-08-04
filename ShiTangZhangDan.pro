#-------------------------------------------------
#
# Project created by QtCreator 2017-08-02T11:02:57
#
#-------------------------------------------------

QT       += core gui
QT+=axcontainer
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ShiTangZhangDan
TEMPLATE = app


SOURCES += main.cpp\
        mainwidget.cpp \
    ExcelBase.cpp \
    QVariantListListModel.cpp

HEADERS  += mainwidget.h \
    ExcelBase.h \
    QVariantListListModel.h

FORMS    += mainwidget.ui

DISTFILES += \
    C:/Users/cx/Desktop/20170621-20170720旧-原始.xls \
    C:/Users/cx/Desktop/20170621-20170720新-原始.xls
