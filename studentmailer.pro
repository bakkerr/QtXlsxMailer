#-------------------------------------------------
#
# Project created by QtCreator 2016-02-29T23:35:22
#
#-------------------------------------------------

QT       += core gui xlsx

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = studentmailer
TEMPLATE = app

SOURCES += main.cpp\
        mainwindow.cpp \
    xlsxsheetmodel.cpp

HEADERS  += mainwindow.h \
    xlsxsheetmodel.h \
    xlsxsheetmodel_p.h

win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../SmtpClient-for-Qt/release/ -lSMTPEmail
else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../SmtpClient-for-Qt/debug/ -lSMTPEmail
else:unix: LIBS += -L$$PWD/../SmtpClient-for-Qt/ -lSMTPEmail

INCLUDEPATH += $$PWD/../SmtpClient-for-Qt/src
DEPENDPATH += $$PWD/../SmtpClient-for-Qt/src
