#-------------------------------------------------
#
# Project created by QtCreator 2015-06-29T09:12:07
#
#-------------------------------------------------


QT       += core gui
QT += axcontainer
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

OBJECTS_DIR += obj
UI_DIR += forms
RCC_DIR += rcc
MOC_DIR += moc
DESTDIR += bin

TARGET = Convertion
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp     

HEADERS  += mainwindow.h

FORMS    += mainwindow.ui

CONFIG -= console
CONFIG += warn_off

RESOURCES += \
    res.qrc

DISTFILES += \
    quit_pressed.png
