#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDockWidget>

#include <QTabWidget>
#include <QTextEdit>
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QLCDNumber>
#include <QLabel>

#include <smtpclient.h>
#include <mimetext.h>

/* Compile-time constant values. */
#define APPLICATION_VERSION       "0.2"
#define APPLICATION_NAME          "Qt XLSX Email Generator"
#define APPLICATION_AUTHOR        "Roy Bakker"
#define APPLICATION_AUTHOR_EMAIL  "baroy@hr.nl"
#define APPLICATION_AUTHOR_URL    "www.roybakker.nl"
#define APPLICATION_COMPANY       "Hogeschool Rotterdam"
#define APPLICATION_YEAR          "2016"
#define APPLICATION_URL           "http://github.com/bakkerr/"

/* MainWindow class. */
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

/* Callback functions. */
private slots:
    /* When new sheet is loaded. */
    void updateSheet();

    /* When fields are is updated. */
    void updateInfo();

    /* When preview should be updated. */
    void updateText();

    /* Load sheet dialog */
    void loadSheet();

    /* Generate mail contents. */
    void addNewTextTab();
    void generateText(bool newTab);
    void generateNewText()     { generateText(true);  }
    void generateReplaceText() { generateText(false); }
    void toggleGenerateWidget(bool s);

    /* Close or rename a tab in the XLSX viewer or the Editor. */
    void closeTab(int index);
    void renameTab(int index);

    /* Update blocker when adding values to comboboxes. */
    void blockRowSignals(bool b);

    /* Handle the SMTP (dis)connect. */
    void SMTPconnect();
    void SMTPdisconnect();
    void toggleSMTPWidget(bool b);

    /* The main thing... Sending mails */
    void sendMails();

    /* Show about dialog. */
    void about();

private:
    /* (Dock)widget generators. */
    void createGeneralOptionsWidget();
    void createSMTPWidget();

    void createEditorWidget();
    void createGenerateWidget();

    void createPreviewWidget();
    void createRowSelectWidget();

    void createXlsxViewerWidget();

    QWidget *createInfoBar();

    /* Valid email address? */
    bool isValidEmail(QString address);

    /* Generate mailtext from template */
    QString getMailText(int offset);

    /* Row and column parser */
    QString getData(QString cell, int offset);

    /* Extract data from spreadsheet. */
    QString getData(int row, int col);

    /* Dockwidgets */
    QDockWidget *m_generalOptionsDW;
    QDockWidget *m_xlsxViewerDW;
    QDockWidget *m_editorDW;
    QDockWidget *m_previewDW;

    /* SMTP client */
    SmtpClient *m_smtp;

    /* General Options fields */
    QLineEdit *m_emailSubject;
    QLineEdit *m_emailBcc;
    QLineEdit *m_senderName;
    QLineEdit *m_senderEmail;
    QLineEdit *m_courseCode;

    /* SMTP settings. */
    QWidget *m_SMTPWidget;
    QPushButton *m_SMTPWidgetHideButton;
    QLineEdit *m_SMTPserver;
    QLineEdit *m_SMTPport;
    QComboBox *m_SMTPtype;

    /* XLSX viewer. */
    QTabWidget *m_xlsxTab;

    /* Editor/Composer. */
    QTabWidget *m_textTab;
    QFrame *m_generateWidget;
    QPushButton *m_generateButton;
    QComboBox *m_nameCol;
    QComboBox *m_finalCol;
    QComboBox *m_startCol;
    QComboBox *m_stopCol;
    QComboBox *m_maxRow;
    QComboBox *m_headerRow;

    /* Selection and Preview. */
    QTextEdit *m_view;
    QFrame *m_rowSelectWidget;
    QComboBox *m_emailColumn;
    QLineEdit *m_emailAppend;
    QComboBox *m_start;
    QComboBox *m_stop;
    QComboBox *m_preview;
    QLCDNumber *m_nMails;

    /* Bottom text */
    QLabel *m_extraText;

};

#endif // MAINWINDOW_H
