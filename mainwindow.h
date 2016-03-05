#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QTabWidget>
#include <QTextEdit>
#include <QLineEdit>
#include <QSpinBox>
#include <QComboBox>
#include <QLCDNumber>

#include <smtpclient.h>
#include <mimetext.h>

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

    void blockUpdate(bool b);

    void sendMails();
    void SMTPconnect();
    void SMTPdisconnect();

private:
    QWidget *createRowSelectWidget();
    QWidget *createGeneralOptionsWidget();
    QWidget *createEditorWidget();
    QWidget *createPreviewWidget();
    QWidget *createXlsxViewer();
    QWidget *createMiddleWidget();

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

    QTabWidget *m_tabWidget;

    QTextEdit *m_edit;
    QTextEdit *m_view;

    QComboBox *m_start;
    QComboBox *m_stop;
    QComboBox *m_preview;
    QLCDNumber *m_nMails;
};

#endif // MAINWINDOW_H
