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
#include <QCheckBox>

#include <QPropertyAnimation>

#include <smtpclient.h>

/* Compile-time constant values. */
#define APPLICATION_VERSION       "0.2"
#define APPLICATION_NAME          "Qt XLSX Email Generator"
#define APPLICATION_NAME_ABBR     "StudentMailer"
#define APPLICATION_AUTHOR        "Roy Bakker"
#define APPLICATION_AUTHOR_EMAIL  "baroy@hr.nl"
#define APPLICATION_AUTHOR_URL    "www.roybakker.nl"
#define APPLICATION_COMPANY       "Hogeschool Rotterdam"
#define APPLICATION_COMPANY_ABBR  "HR"
#define APPLICATION_YEAR          "2016"
#define APPLICATION_URL           "http://github.com/bakkerr/"

/* Debugging. */
#define DO_NOT_SEND_EMAILS 0

/* MainWindow class. */
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

signals:

/* Callback functions. */
private slots:

    /* [3] Save/load/delete Settings. */
    void saveSettings();
    void loadSettings();
    void deleteSettings();
    void closeEvent(QCloseEvent *closeEvent);

    /* [4] Show/hide widgets. */
    void toggleSettingsWidget(bool b);
    void toggleSMTPWidget(bool b);
    void toggleGenerateWidget(bool s);
    void toggleAttachmentWidget(bool s);

    /*
     * [5] Update / load.
     */

    /* Add attachment. */
    void addAttachment();

    /* Delete Attachment. */
    void deleteAttachment();

    void selectAttachmentDirectory();

    /* Load sheet dialog */
    void loadSheet();

    /* Slot called when selecting an onther sheet. */
    void updateSheet();

    /* When fields are updated. */
    void updateInfo();

    /* When preview should be updated. */
    void updateText();

    /* Update blocker when adding values to comboboxes. */
    void blockRowSignals(bool b);

    /* [6] Generate mail contents. */
    void addNewTextTab();
    void generateText(bool newTab);
    void generateNewText()     { generateText(true);  }
    void generateReplaceText() { generateText(false); }

    /* [7] Close or rename a tab in the XLSX viewer or the Editor. */
    void closeTab(int index);
    void renameTab(int index);

    /*
     * [8] SMTP.
     */

    /* Handle the SMTP (dis)connect. */
    void SMTPconnect();
    void SMTPdisconnect();

    /* The main thing... Sending mails */
    void sendMails();
    bool sendMail(MimeMessage *m);

    /* [9] Show about dialog. */
    void about();

private:

    /*
     * [1] UI generators.
     */

    /* Settings Dockwidget */
    void createGeneralOptionsWidget();
    void createSettingsWidget();
    void createSMTPWidget();

    /* Editor Dockwidget */
    void createEditorWidget();
    void createGenerateWidget();

    /* Preview Dockwidget */
    void createPreviewWidget();
    void createMailSelectWidget();
    void createAttachmentWidget();

    /* Xlsx Viewer Dockwidget. */
    void createXlsxViewerWidget();

    /* Toolbar */
    void createToolBar();

    /*
     * [2] General methods.
     */

    /* Generate mailtext or header from template */
    QString getMailHeader(int offset);
    QString getMailText(int offset);

    /* Row and column parser */
    QString getData(QString cell, int offset);

    /* Extract data from spreadsheet. */
    QString getData(int row, int col);

    /* Valid email address? */
    bool isValidEmail(QString address);
    bool isValidHRStudentEmail(QString address);
    bool isValidHREmployeeEmail(QString address);

    /*
     * Private members.
     */

    /* Dockwidgets */
    QDockWidget *m_generalOptionsDW;
    QDockWidget *m_xlsxViewerDW;
    QDockWidget *m_editorDW;
    QDockWidget *m_previewDW;

    /* SMTP client */
    SmtpClient *m_SMTPConnection;

    /* General Options fields */
    QLineEdit *m_emailSubject;
    QLineEdit *m_emailBcc;
    QLineEdit *m_reportCC;
    QLineEdit *m_senderName;
    QLineEdit *m_senderEmail;
    QLineEdit *m_courseCode;
    QPushButton *m_addAttachment;
    QPushButton *m_deleteSelectedAttachment;
    QComboBox *m_attachments;

    /* Settings settings. */
    QFrame *m_settingsWidget;
    QPropertyAnimation *m_toggleSettingsAnimation;
    QPushButton *m_settingsWidgetToggleButton;
    QCheckBox *m_runtimeValidate;
    QCheckBox *m_saveOnExitCheckBox;
    QCheckBox *m_validateHR;

    /* SMTP settings. */
    QFrame *m_SMTPWidget;
    QPropertyAnimation *m_toggleSMTPAnimation;
    QPushButton *m_SMTPWidgetToggleButton;
    QLineEdit *m_SMTPserver;
    QLineEdit *m_SMTPport;
    QComboBox *m_SMTPtype;

    /* XLSX viewer. */
    QToolButton *m_loadXlsxFileButton;
    QTabWidget *m_xlsxTab;

    /* Editor/Composer. */
    QTabWidget *m_textTab;
    QFrame *m_generateWidget;
    QPropertyAnimation *m_toggleGenerateAnimation;
    QPushButton *m_generateWidgetToggleButton;
    QComboBox *m_nameColSelect;
    QComboBox *m_finalGradeColSelect;
    QComboBox *m_startColSelect;
    QComboBox *m_stopColSelect;
    QComboBox *m_maxRowSelect;
    QComboBox *m_headerRowSelect;

    /* Selection and Preview. */
    QTextEdit *m_previewText;
    QFrame *m_rowSelectWidget;
    QComboBox *m_emailColumnSelect;
    QLineEdit *m_emailAppendText;
    QComboBox *m_firstRowSelect;
    QComboBox *m_lastRowSelect;
    QComboBox *m_previewSelect;
    QLCDNumber *m_nMailsDisplay;

    /* Attacment Widget */
    QFrame *m_attachmentWidget;
    QPropertyAnimation *m_toggleAttachmentAnimation;
    QPushButton *m_attachmentWidgetToggleButton;
    QString m_attachmentDirectory;
    QLineEdit *m_attachmentAppend;
    QPushButton *m_selectAttachmentDirectoryButton;
    QComboBox *m_attachmentColSelect;

    /* Toolbar */
    QToolBar *m_toolBar;

};

#endif // MAINWINDOW_H
