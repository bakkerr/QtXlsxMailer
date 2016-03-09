#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QTabWidget>
#include <QTextEdit>
#include <QLineEdit>
#include <QComboBox>
#include <QLCDNumber>
#include <QLabel>

#include <smtpclient.h>
#include <mimetext.h>

#define APPLICATION_VERSION       "0.1"
#define APPLICATION_NAME          "Qt XLSX Email Generator"
#define APPLICATION_AUTHOR        "Roy Bakker"
#define APPLICATION_AUTHOR_EMAIL  "baroy@hr.nl"
#define APPLICATION_AUTHOR_URL    "www.roybakker.nl"
#define APPLICATION_COMPANY       "Hogeschool Rotterdam"
#define APPLICATION_YEAR          "2016"
#define APPLICATION_URL           "http://github.com/bakkerr/"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void updateSheet();
    void updateInfo();
    void updateText();
    void loadSheet();
    void generateTextDialog();

    void closeTab(int index);

    void blockRowSignals(bool b);

    void sendMails();
    void SMTPconnect();
    void SMTPdisconnect();

    void about();

private:
    QWidget *createGeneralOptionsWidget();
    QWidget *createEditorWidget();
    QWidget *createRowSelectWidget();
    QWidget *createPreviewWidget();
    QWidget *createMiddleWidget();
    QWidget *createXlsxViewer();
    QWidget *createInfoBar();


    QString getMailText(int offset);
    QString getData(QString cell, int offset);
    QString getData(int row, int column);

    SmtpClient *m_smtp;

    QComboBox *m_emailColumn;
    QLineEdit *m_emailAppend;
    QLineEdit *m_emailSubject;
    QLineEdit *m_emailBcc;
    QLineEdit *m_senderName;
    QLineEdit *m_senderEmail;
    QLineEdit *m_courseCode;

    QLineEdit *m_SMTPserver;
    QLineEdit *m_SMTPport;
    QComboBox *m_SMTPtype;

    QTabWidget *m_tabWidget;

    QTextEdit *m_edit;
    QTextEdit *m_view;

    QComboBox *m_start;
    QComboBox *m_stop;
    QComboBox *m_preview;
    QLCDNumber *m_nMails;

    QLabel *m_extraText;
};

#endif // MAINWINDOW_H
