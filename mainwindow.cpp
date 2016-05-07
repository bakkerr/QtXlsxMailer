#include "mainwindow.h"

#include <QApplication>

#include <QWidget>
#include <QTableView>

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QGridLayout>

#include <QFileDialog>
#include <QInputDialog>
#include <QMessageBox>
#include <QProgressBar>

#include <QToolBar>
#include <QToolButton>
#include <QAction>

#include <QRegExp>
#include <QStringRef>

#include <mimetext.h>
#include <mimeattachment.h>

#include <QtXlsx>
#include "xlsxsheetmodel.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent)
{

    this->setWindowTitle(tr("Qt XLSX Email Generator [Hogeschool Rotterdam]"));

    /* Make sure the SMTP connection pointer is NULL. */
    m_SMTPConnection = NULL;

    /* Dockwidgets options. */
    setDockNestingEnabled(true);
    setAnimated(true);

    /* Create Dockwidgets and toolbar. */
    createGeneralOptionsWidget();
    createEditorWidget();
    createPreviewWidget();
    createXlsxViewerWidget();
    createToolBar();

    /* Add Dockwidgets and toolbar. */
    this->addDockWidget(Qt::TopDockWidgetArea, m_generalOptionsDW);
    this->addDockWidget(Qt::LeftDockWidgetArea, m_editorDW);
    this->addDockWidget(Qt::RightDockWidgetArea, m_previewDW);
    this->addDockWidget(Qt::BottomDockWidgetArea, m_xlsxViewerDW);
    this->addToolBar(Qt::BottomToolBarArea, m_toolBar);

    /* Having no central widget gives problems in layout. */
    QWidget *cw = new QWidget(this);
    cw->setFixedSize(0, 0);
    setCentralWidget(cw);

    /* Set default values. */
    loadSettings();
    updateSheet();

}

/*
 * Destructor.
 *
 * TODO: clean-up properly?
 */
MainWindow::~MainWindow(){

}

/*
 * [1] UI generators.
 */

/* General Options Dockwidget */
void MainWindow::createGeneralOptionsWidget(){

    m_generalOptionsDW = new QDockWidget(tr("General parameters"), this);
    m_generalOptionsDW->setFixedHeight(175);

    QFrame *generalOptionsWidget = new QFrame(m_generalOptionsDW);
    generalOptionsWidget->setFrameShape(QFrame::StyledPanel);

    QGridLayout *generalOptionsLayout = new QGridLayout(generalOptionsWidget);
    QMargins m = generalOptionsLayout->contentsMargins();
    m.setLeft(0);
    m.setRight(0);
    generalOptionsLayout->setContentsMargins(m);

    /* Sender name field */
    m_senderName = new QLineEdit(tr(""), generalOptionsWidget);
    m_senderName->setToolTip(tr("This field contains the name of the\n"
                                "sender of the emails. It will also be\n"
                                "used in the automatic generation of\n"
                                "the email text.\n\n"
                                "Example: \"D. Ocent\" or \"Do Cent\""));
    connect(m_senderName, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* Sender email address field. */
    m_senderEmail = new QLineEdit(tr(""), generalOptionsWidget);
    m_senderEmail->setToolTip(tr("This field will be used as the sender\n"
                                 "email address of the mails that are sent.\n"
                                 "It will also be used in the SMTP connection\n"
                                 "as the SMTP username.\n\n"
                                 "Example: \"docentcode@hr.nl\""));
    m_senderEmail->setValidator(new QRegExpValidator(QRegExp("[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63}", Qt::CaseInsensitive), m_senderEmail));
    connect(m_senderEmail, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* Email subject field. */
    m_emailSubject = new QLineEdit(tr(""), generalOptionsWidget);
    m_emailSubject->setToolTip(tr("Add the subject of the generated emails here.\n\n"
                                  "Example: \"Cijfers Tentamen\""));
    connect(m_emailSubject, SIGNAL(textChanged(QString)), this, SLOT(updateText()));
    m_courseCode = new QLineEdit(tr(""), generalOptionsWidget);
    m_courseCode->setToolTip(tr("Add the coursecode here.\n"
                                "It will be added as a [tag] to the\n"
                                "subject of the generated emails.\n\n"
                                "Example: \"ELEVAK01\""));
    connect(m_courseCode, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* Email bcc field. */
    m_emailBcc = new QLineEdit(tr(""), generalOptionsWidget);
    m_emailBcc->setToolTip(tr("Send a (blind) copy of every email to this address.\n"
                              "Multiple addresses may be added seperated by a ';'\n"
                              "and no spaces.\n\n"
                              "Example \"collegue@hr.nl;other@extern.com\""));
    m_emailBcc->setValidator(new QRegExpValidator(QRegExp("(([A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63})[;])*", Qt::CaseInsensitive), m_emailBcc));
    connect(m_emailBcc, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* Report cc field. */
    m_reportCC = new QLineEdit(tr(""), generalOptionsWidget);
    m_reportCC->setToolTip(tr("Send a copy of the report to this address.\n"
                              "Multiple addresses may be added seperated by a ';'\n"
                              "and no spaces.\n\n"
                              "Example \"collegue@hr.nl;other@extern.com\""));
    m_reportCC->setValidator(new QRegExpValidator(QRegExp("(([A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63})[;])*", Qt::CaseInsensitive), m_reportCC));
    connect(m_reportCC, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    m_attachments = new QComboBox(generalOptionsWidget);
    m_attachments->setToolTip(tr("These attachments will be added to all emails."));
    m_attachments->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::MinimumExpanding);
    m_attachments->hide();

    m_addAttachment = new QPushButton(tr("+"), generalOptionsWidget);
    m_addAttachment->setToolTip(tr("Add attachment."));
    m_addAttachment->setMaximumWidth(30);
    connect(m_addAttachment, SIGNAL(clicked()), this, SLOT(addAttachment()));

    m_deleteSelectedAttachment = new QPushButton(tr("-"), generalOptionsWidget);
    m_deleteSelectedAttachment->setToolTip(tr("Delete selected attachment."));
    m_deleteSelectedAttachment->setMaximumWidth(30);
    connect(m_deleteSelectedAttachment, SIGNAL(clicked()), this, SLOT(deleteAttachment()));
    m_deleteSelectedAttachment->hide();

    QHBoxLayout *attachmentLayout = new QHBoxLayout();
    attachmentLayout->addWidget(m_addAttachment);
    attachmentLayout->addWidget(m_deleteSelectedAttachment);
    attachmentLayout->addWidget(m_attachments);

    /* Create the SMTP settings widget. */
    createSMTPWidget();

    createSettingsWidget();

    /* Add all fields to the layout. */
    generalOptionsLayout->addWidget(m_settingsWidgetToggleButton, 0, 1, 4, 1);
    generalOptionsLayout->addWidget(m_settingsWidget, 0, 0, 4, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Name:"), generalOptionsWidget), 0, 2);
    generalOptionsLayout->addWidget(m_senderName, 0, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Email:"), generalOptionsWidget), 1, 2);
    generalOptionsLayout->addWidget(m_senderEmail, 1, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Add BCC:"), generalOptionsWidget), 2, 2);
    generalOptionsLayout->addWidget(m_emailBcc, 2, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Report CC:"), generalOptionsWidget), 3, 2);
    generalOptionsLayout->addWidget(m_reportCC, 3, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Email Subject:"), generalOptionsWidget), 0, 4);
    generalOptionsLayout->addWidget(m_emailSubject, 0, 5);
    generalOptionsLayout->addWidget(new QLabel(tr("Course Code:"), generalOptionsWidget), 1, 4);
    generalOptionsLayout->addWidget(m_courseCode, 1, 5);
    generalOptionsLayout->addWidget(new QLabel(tr("Attachments:"), generalOptionsWidget), 2, 4);
    generalOptionsLayout->addLayout(attachmentLayout, 2, 5);
    generalOptionsLayout->setAlignment(attachmentLayout, Qt::AlignLeft);
    generalOptionsLayout->addWidget(m_SMTPWidgetToggleButton, 0, 6, 4, 1);
    generalOptionsLayout->addWidget(m_SMTPWidget, 0, 7, 4, 1);

    /* Set layout to mainwidget. */
    generalOptionsWidget->setLayout(generalOptionsLayout);

    /* Set mainwidget on DockWidget. */
    m_generalOptionsDW->setWidget(generalOptionsWidget);

}

/* Settings Widget (can hide). */
void MainWindow::createSettingsWidget(){
    m_settingsWidget = new QFrame(m_generalOptionsDW);

    QGridLayout *settingsLayout = new QGridLayout(m_settingsWidget);
    QMargins m = settingsLayout->contentsMargins();
    m.setTop(0);
    m.setBottom(0);
    m.setRight(0);
    settingsLayout->setContentsMargins(m);

    m_runtimeValidate = new QCheckBox(tr("Direct Validation"), m_settingsWidget);
    m_runtimeValidate->setToolTip(tr("Color boxes and buttons red when they have invalid content as you type."));
    m_runtimeValidate->setChecked(true);
    connect(m_runtimeValidate, SIGNAL(stateChanged(int)), this, SLOT(updateText()));
    m_validateHR = new QCheckBox(tr("Validate for HR"), m_settingsWidget);
    m_validateHR->setToolTip(tr("Validate student email addresses ([7 digits]@hr.nl) and employee code ([5 characters]@hr.nl) for use at the Hogeschool Rotterdam."));
    m_validateHR->setChecked(true);
    m_saveOnExitCheckBox = new QCheckBox(tr("Ask to save on exit"), m_settingsWidget);
    m_saveOnExitCheckBox->setToolTip(tr("Ask before saving text tabs and general parameters on exit. When not checked, these values are automatically saved."));
    m_saveOnExitCheckBox->setChecked(false);

    m_toggleSettingsAnimation = new QPropertyAnimation(m_settingsWidget, "maximumWidth");
    m_toggleSettingsAnimation->setDuration(500);
    m_settingsWidgetToggleButton = new QPushButton(tr("<"), m_generalOptionsDW);
    m_settingsWidgetToggleButton->setMaximumWidth(20);
    m_settingsWidgetToggleButton->setMinimumWidth(20);
    m_settingsWidgetToggleButton->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);
    connect(m_settingsWidgetToggleButton, SIGNAL(toggled(bool)), this, SLOT(toggleSettingsWidget(bool)));
    connect(m_toggleSettingsAnimation, SIGNAL(finished()), this, SLOT(repaint()));
    m_settingsWidgetToggleButton->setCheckable(true);
    m_settingsWidgetToggleButton->setChecked(true);
    m_settingsWidgetToggleButton->setChecked(false);

    /* Clear settings button. */
    QPushButton *clearSettingsButton = new QPushButton(tr("Clear Saved Settings"), m_settingsWidget);
    clearSettingsButton->setToolTip(tr("Reset all saved and edited settings.\nThe application will be as new."));
    connect(clearSettingsButton, SIGNAL(clicked()), this, SLOT(deleteSettings()));

    /* Save settings button. */
    QPushButton *saveSettingsButton = new QPushButton(tr("Save Settings"), m_settingsWidget);
    saveSettingsButton->setToolTip(tr("Save settings now."));
    connect(saveSettingsButton, SIGNAL(clicked()), this, SLOT(saveSettings()));

    /* Set Layout. */
    settingsLayout->addWidget(saveSettingsButton, 0, 1);
    settingsLayout->addWidget(clearSettingsButton, 1, 1);
    settingsLayout->addWidget(m_runtimeValidate, 0, 2);
    settingsLayout->addWidget(m_validateHR, 1, 2);
    settingsLayout->addWidget(m_saveOnExitCheckBox, 2, 2);

}

/* Creates the SMTP settings widget (can hide). */
void MainWindow::createSMTPWidget(){

    m_SMTPWidget = new QFrame(m_generalOptionsDW);

    QGridLayout *smtpSettingsLayout = new QGridLayout(m_SMTPWidget);
    QMargins m = smtpSettingsLayout->contentsMargins();
    m.setBottom(0);
    m.setTop(0);
    m.setLeft(0);
    smtpSettingsLayout->setContentsMargins(m);

    m_SMTPserver = new QLineEdit(tr(""), m_SMTPWidget);

    m_SMTPport = new QLineEdit(tr(""), m_SMTPWidget);
    m_SMTPport->setValidator(new QIntValidator(1, 65535, m_SMTPport));

    m_SMTPtype = new QComboBox(m_SMTPWidget);
    m_SMTPtype->addItem(tr("SSL"), SmtpClient::SslConnection);
    m_SMTPtype->addItem(tr("TLS"), SmtpClient::TlsConnection);
    m_SMTPtype->addItem(tr("TCP"), SmtpClient::TcpConnection);

    QPushButton *SMTPConnectButton = new QPushButton(tr("SMTP Connect"), m_SMTPWidget);
    SMTPConnectButton->setToolTip(tr("Connect to the SMTP server now."));
    connect(SMTPConnectButton, SIGNAL(clicked()), this, SLOT(SMTPconnect()));

    m_toggleSMTPAnimation = new QPropertyAnimation(m_SMTPWidget, "maximumWidth");
    m_toggleSMTPAnimation->setDuration(500);
    m_SMTPWidgetToggleButton = new QPushButton(tr("<"), m_generalOptionsDW);
    m_SMTPWidgetToggleButton->setMaximumWidth(20);
    m_SMTPWidgetToggleButton->setMinimumWidth(20);
    m_SMTPWidgetToggleButton->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);
    connect(m_SMTPWidgetToggleButton, SIGNAL(toggled(bool)), this, SLOT(toggleSMTPWidget(bool)));
    connect(m_toggleSMTPAnimation, SIGNAL(finished()), this, SLOT(repaint()));
    m_SMTPWidgetToggleButton->setCheckable(true);
    m_SMTPWidgetToggleButton->setChecked(true);
    m_SMTPWidgetToggleButton->setChecked(false);

    smtpSettingsLayout->addWidget(new QLabel(tr("SMTP server:"), m_SMTPWidget), 0, 1);
    smtpSettingsLayout->addWidget(m_SMTPserver, 0, 2);
    smtpSettingsLayout->addWidget(new QLabel(tr("SMTP port:"), m_SMTPWidget), 1, 1);
    smtpSettingsLayout->addWidget(m_SMTPport, 1, 2);
    smtpSettingsLayout->addWidget(m_SMTPtype, 2, 1);
    smtpSettingsLayout->addWidget(SMTPConnectButton, 2, 2);

    m_SMTPWidget->setLayout(smtpSettingsLayout);

}

/* Create Editor Dockwidget */
void MainWindow::createEditorWidget(){

    /* Create Dockwidget for the editor. */
    m_editorDW = new QDockWidget(tr("Edit:"), this);
    m_editorDW->setMinimumHeight(375);

    /* Create main widget. */
    QFrame *editorWidget = new QFrame(m_editorDW);
    editorWidget->setFrameShape(QFrame::StyledPanel);

    /* Create Layout. */
    QHBoxLayout *editorWidgetLayout = new QHBoxLayout(editorWidget);

    /* Create tabwidget to contain editors. */
    m_textTab = new QTabWidget(editorWidget);
    m_textTab->setToolTip(tr("The selected tab will be used to\n"
                             "generate the e-mails from.\n\n"
                             "Double-click to rename tab."));
    m_textTab->setTabsClosable(true);

    /* Button to add new tab. */
    QToolButton *addTabButton = new QToolButton(m_textTab);
    addTabButton->setText(tr("+"));
    addTabButton->setToolTip(tr("Add new tab."));
    connect(addTabButton, SIGNAL(clicked()), this, SLOT(addNewTextTab()));
    m_textTab->addTab(new QWidget(m_textTab), tr(""));
    m_textTab->setTabEnabled(0, false);
    m_textTab->tabBar()->setTabButton(0, QTabBar::RightSide, addTabButton);

    /* Connect signals for close, update and rename. */
    connect(m_textTab, SIGNAL(tabCloseRequested(int)), this, SLOT(closeTab(int)));
    connect(m_textTab, SIGNAL(currentChanged(int)), this, SLOT(updateText()));
    connect(m_textTab, SIGNAL(tabBarDoubleClicked(int)), this, SLOT(renameTab(int)));

    /* Generate the sliding text generator widget. */
    createGenerateWidget();

    /* Add widgets to layout */
    editorWidgetLayout->addWidget(m_textTab);
    editorWidgetLayout->addWidget(m_generateWidgetToggleButton);
    editorWidgetLayout->addWidget(m_generateWidget);

    /* Set layout to main widget. */
    editorWidget->setLayout(editorWidgetLayout);

    /* Set main widget to dockwidget. */
    m_editorDW->setWidget(editorWidget);

}

/* Create widget to generate email text. */
void MainWindow::createGenerateWidget(){

    /* Create a frame for this widget. */
    m_generateWidget = new QFrame(m_editorDW);
    m_generateWidget->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);

    /* Create Layout. */
    QGridLayout *generateWidgetLayout = new QGridLayout(m_generateWidget);
    generateWidgetLayout->setContentsMargins(0, 0, 0, 0);

    /* Create Widgets */
    m_nameColSelect = new QComboBox(m_generateWidget);
    m_nameColSelect->setToolTip(tr("The column to use for the name of the recipient.\n"
                             "Select <none> if you do not want to include this."));
    m_finalGradeColSelect = new QComboBox(m_generateWidget);
    m_finalGradeColSelect->setToolTip(tr("The column to use for the final grade.\n"
                              "Select <none> if you do not want to include this."));
    m_startColSelect = new QComboBox(m_generateWidget);
    m_startColSelect->setToolTip(tr("We can include a selection of columns to include.\n"
                              "Specify the column to start with here.\n"
                              "Select <none> if you do not want to include this."));
    m_stopColSelect = new QComboBox(m_generateWidget);
    m_stopColSelect->setToolTip(tr("We can include a selection of columns to include.\n"
                             "Specify the column to end with here.\n"
                             "Select <none> if you do not want to include this."));
    m_maxRowSelect = new QComboBox(m_generateWidget);
    m_maxRowSelect->setToolTip(tr("We can include the maximum score or default value for\n"
                            "the columns you have selected.\n"
                            "Specify the row to use for this here.\n"
                            "Select <none> if you do not want to include this."));
    m_headerRowSelect = new QComboBox(m_generateWidget);
    m_headerRowSelect->setToolTip(tr("We can include names for the columns you have selected.\n"
                             "Specify the row to use for this here.\n"
                             "Select <none> if you do not want to include this."));

    /* Button to create text in current tab. */
    QPushButton *replaceButton = new QPushButton(tr("Current Tab"), m_generateWidget);
    replaceButton->setToolTip(tr("Generate text and overwrite selected tab."));
    connect(replaceButton, SIGNAL(clicked()), this, SLOT(generateReplaceText()));

    /* Button to create text in new tab. */
    QPushButton *newButton = new QPushButton(tr("New Tab"), m_generateWidget);
    newButton->setToolTip(tr("Generate text and place in new tab."));
    connect(newButton, SIGNAL(clicked()), this, SLOT(generateNewText()));

    /* Define animation for adjusting the maximumWidth of the widget. */
    m_toggleGenerateAnimation = new QPropertyAnimation(m_generateWidget, "maximumWidth");
    m_toggleGenerateAnimation->setDuration(500);

    /* Button to show/hide this widget. */
    m_generateWidgetToggleButton = new QPushButton(tr("<"), m_editorDW);
    m_generateWidgetToggleButton->setMinimumWidth(20);
    m_generateWidgetToggleButton->setMaximumWidth(20);
    m_generateWidgetToggleButton->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);
    m_generateWidgetToggleButton->setCheckable(true);
    connect(m_generateWidgetToggleButton, SIGNAL(toggled(bool)), this, SLOT(toggleGenerateWidget(bool)));
    connect(m_toggleGenerateAnimation, SIGNAL(finished()), this, SLOT(repaint()));
    m_generateWidgetToggleButton->setChecked(true);

    /* Add labels to layout */
    generateWidgetLayout->addWidget(new QLabel(tr("Generate email text:"), m_generateWidget), 0, 0, 1, 2);
    generateWidgetLayout->addWidget(new QLabel(tr("Name [col]:"), m_generateWidget), 1, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("Final grade [col]:"), m_generateWidget), 2, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("Include [col]:"), m_generateWidget), 3, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("up to (incl) [col]:"), m_generateWidget), 4, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("Grade title [row]:"), m_generateWidget), 5, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("Max. points [row]:"), m_generateWidget), 6, 0);
    generateWidgetLayout->addWidget(new QLabel(tr("Generate text in:"), m_generateWidget), 8, 0);

    /* Add widgets to layout */
    generateWidgetLayout->addWidget(m_nameColSelect, 1, 1);
    generateWidgetLayout->addWidget(m_finalGradeColSelect, 2, 1);
    generateWidgetLayout->addWidget(m_startColSelect, 3, 1);
    generateWidgetLayout->addWidget(m_stopColSelect, 4, 1);
    generateWidgetLayout->addWidget(m_headerRowSelect, 5, 1);
    generateWidgetLayout->addWidget(m_maxRowSelect, 6, 1);
    generateWidgetLayout->setRowStretch(7, 40);
    generateWidgetLayout->addWidget(newButton, 9, 0);
    generateWidgetLayout->addWidget(replaceButton, 9, 1);

    /* Set layout to main widget. */
    generateWidgetLayout->setAlignment(Qt::AlignTop);
    m_generateWidget->setLayout(generateWidgetLayout);

}

/* Create the selection and preview widget. */
void MainWindow::createPreviewWidget(){

    /* Create the dockwidget. */
    m_previewDW = new QDockWidget(tr("Selection and preview:"), this);
    m_previewDW->setMinimumHeight(375);

    /* Create a (main)frame for this widget. */
    QFrame *previewWidget = new QFrame(m_previewDW);
    previewWidget->setFrameShape(QFrame::StyledPanel);

    /* Create main layout. */
    QHBoxLayout *previewWidgetLayout = new QHBoxLayout();
    previewWidgetLayout->setAlignment(Qt::AlignTop);

    /* Create layouts. */
    QVBoxLayout *previewBoxLayout = new QVBoxLayout();
    QHBoxLayout *previewSelectionLayout = new QHBoxLayout();

    /* Create 7-segment display for the number of mails. */
    m_nMailsDisplay = new QLCDNumber(4, m_previewDW);
    m_nMailsDisplay->setFrameStyle(QFrame::NoFrame);
    m_nMailsDisplay->setSegmentStyle(QLCDNumber::Filled);
    m_nMailsDisplay->setPalette(QPalette(Qt::red));
    m_nMailsDisplay->setToolTip(tr("This is the number of emails\nthat this program will send."));

    /* Selection for the row (email) to preview. */
    m_previewSelect = new QComboBox(m_previewDW);
    m_previewSelect->setToolTip(tr("Select the row (email) you want to preview."));
    connect(m_previewSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    /* The preview tool itself is a read-only textbox. */
    m_previewText = new QTextEdit(m_previewDW);
    m_previewText->setToolTip(tr("This is how the email looks as it will be send.\n"
                          "You can select an other email in the selection box above."));
    m_previewText->setReadOnly(true);

    /*
     * Create the widget to select the rows to use and the columns for
     * the email address. Optionally you can add a value to the email address.
     */
    createMailSelectWidget();

    /* Set it all in the layouts. */
    previewSelectionLayout->addWidget(new QLabel(tr("Preview:"), m_previewDW));
    previewSelectionLayout->addWidget(m_previewSelect);
    previewSelectionLayout->addWidget(m_nMailsDisplay);
    previewBoxLayout->addLayout(previewSelectionLayout);
    previewBoxLayout->addWidget(m_previewText);

    previewWidgetLayout->addWidget(m_rowSelectWidget);
    previewWidgetLayout->addLayout(previewBoxLayout);

    /* Set layout to main widget. */
    previewWidget->setLayout(previewWidgetLayout);

    /* Set main widget in the dockwidget. */
    m_previewDW->setWidget(previewWidget);

}

/* Create widget to select the mail rows. */
void MainWindow::createMailSelectWidget(){

    /* Create Frame. */
    m_rowSelectWidget = new QFrame(this);
    m_rowSelectWidget->setFixedWidth(150);
    m_rowSelectWidget->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);

    /* Create layout. */
    QVBoxLayout *rowSelectLayout = new QVBoxLayout(m_rowSelectWidget);
    rowSelectLayout->setContentsMargins(0, 0, 0, 0);

    /* Create row selection. */
    m_firstRowSelect = new QComboBox(m_rowSelectWidget);
    m_firstRowSelect->setToolTip(tr("Select the row where the first email\nshould be generated from."));
    m_lastRowSelect = new QComboBox(m_rowSelectWidget);
    m_lastRowSelect->setToolTip(tr("Select the row where the last email\nshould be generated from."));
    connect(m_firstRowSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));
    connect(m_lastRowSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    /* Select the column where the email addresses are in. */
    m_emailColumnSelect = new QComboBox(m_rowSelectWidget);
    m_emailColumnSelect->setToolTip(tr("Select the column for the\nemail address to use.\n"
                                       "Note: this column in the spreadhseet\nshould be marked as text,\n"
                                       "not as a number."));
    connect(m_emailColumnSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateText()));

    /* Option to append a value to the addresses in the spreadsheet. */
    m_emailAppendText = new QLineEdit(tr("@hr.nl"), m_rowSelectWidget);
    m_emailAppendText->setToolTip(tr("Add a value that should be appended to\n"
                                     "the column where the email address is in.\n"
                                     "If this column already contains a complete\n"
                                     "email address, this field should be empty."));
    connect(m_emailAppendText, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* The button to send the mails. */
    QPushButton *sendMailsButton = new QPushButton(tr("Send mails"), m_rowSelectWidget);
    sendMailsButton->setToolTip(tr("Pressing this button will check if everything is OK.\n\n"
                                   "If not OK, it will display an error message.\n\n"
                                   "If OK, it will connect to the SMTP server if there \n"
                                   "is no connection yet and tries to send the e-mails.\n\n"
                                   "Finally, a message will be displayed with the result."));
    connect(sendMailsButton, SIGNAL(clicked()), this, SLOT(sendMails()));

    /* Create Attachment Widget. */
    createAttachmentWidget();

    /* Add to layout. */
    rowSelectLayout->addWidget(new QLabel(tr("First mail [row]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_firstRowSelect);
    rowSelectLayout->addWidget(new QLabel(tr("Last mail [row]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_lastRowSelect);
    rowSelectLayout->addWidget(new QLabel(tr("Email address [col]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_emailColumnSelect);
    rowSelectLayout->addWidget(new QLabel(tr("and append:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_emailAppendText);
    rowSelectLayout->addSpacerItem(new QSpacerItem(0, 0, QSizePolicy::Expanding, QSizePolicy::Expanding));
    rowSelectLayout->addWidget(m_attachmentWidgetToggleButton);
    rowSelectLayout->addWidget(m_attachmentWidget);
    rowSelectLayout->addWidget(sendMailsButton);

    /* Set layout. */
    m_rowSelectWidget->setLayout(rowSelectLayout);

}

void MainWindow::createAttachmentWidget(){

    /* Create a frame for this widget. */
    m_attachmentWidget = new QFrame(m_previewDW);
    m_attachmentWidget->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::Maximum);

    QVBoxLayout *attachmentWidgetLayout = new QVBoxLayout(m_attachmentWidget);
    attachmentWidgetLayout->setContentsMargins(0, 0, 0, 0);

    /* Define animation for adjusting the maximumHeight of the widget. */
    m_toggleAttachmentAnimation = new QPropertyAnimation(m_attachmentWidget, "maximumHeight");
    m_toggleAttachmentAnimation->setDuration(500);

    /* Button to show/hide this widget. */
    m_attachmentWidgetToggleButton = new QPushButton(tr("^"), m_previewDW);
    m_attachmentWidgetToggleButton->setMinimumHeight(20);
    m_attachmentWidgetToggleButton->setMaximumHeight(20);
    m_attachmentWidgetToggleButton->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Maximum);
    m_attachmentWidgetToggleButton->setCheckable(true);
    connect(m_attachmentWidgetToggleButton, SIGNAL(toggled(bool)), this, SLOT(toggleAttachmentWidget(bool)));
    connect(m_toggleAttachmentAnimation, SIGNAL(finished()), this, SLOT(repaint()));
    m_attachmentWidgetToggleButton->setChecked(true);
    m_attachmentWidgetToggleButton->setChecked(false);

    m_selectAttachmentDirectoryButton = new QPushButton(tr("Select Directory"), m_attachmentWidget);
    m_selectAttachmentDirectoryButton->setToolTip(tr("Select the directory where the\n"
                                                     "individual attachments are located."));
    connect(m_selectAttachmentDirectoryButton, SIGNAL(clicked()), this, SLOT(selectAttachmentDirectory()));

    m_attachmentColSelect = new QComboBox(m_attachmentWidget);
    m_attachmentColSelect->setToolTip(tr("Select the column to load the\n"
                                         "individual attachment from."));
    connect(m_attachmentColSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    m_attachmentAppend = new QLineEdit(tr(".pdf"), m_attachmentWidget);
    m_attachmentAppend->setToolTip(tr("Text or extension to add to the filename."));
    connect(m_attachmentAppend, SIGNAL(textChanged(QString)), this, SLOT(updateInfo()));

    attachmentWidgetLayout->addWidget(m_selectAttachmentDirectoryButton);
    attachmentWidgetLayout->addWidget(m_attachmentColSelect);
    attachmentWidgetLayout->addWidget(m_attachmentAppend);

    m_attachmentWidget->setLayout(attachmentWidgetLayout);
}

/* Create the preview tool for xlsx sheets. */
void MainWindow::createXlsxViewerWidget(){

    /* Create the dockwidget that uses all of the available (remaining) space. */
    m_xlsxViewerDW = new QDockWidget(tr("XLSX Viewer:"), this);
    m_xlsxViewerDW->setMinimumHeight(200);
    m_xlsxViewerDW->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::MinimumExpanding);

    /* Create a frame for this widget that uses all of the available (remaining) space. */
    QFrame *xlsxWidget = new QFrame(m_xlsxViewerDW);
    xlsxWidget->setFrameShape(QFrame::StyledPanel);

    QVBoxLayout *xlsxWidgetLayout = new QVBoxLayout();

    /* Add the tabwidget where the tabs from the xlsx file can be loaded. */
    m_xlsxTab = new QTabWidget(m_xlsxViewerDW);
    m_xlsxTab->setTabsClosable(true);
    m_xlsxTab->setToolTip(tr("The selected tab will be used to\n"
                             "generate the e-mails from."));
    connect(m_xlsxTab, SIGNAL(tabCloseRequested(int)), this, SLOT(closeTab(int)));
    connect(m_xlsxTab, SIGNAL(currentChanged(int)), this, SLOT(updateSheet()));
    connect(m_xlsxTab, SIGNAL(tabBarDoubleClicked(int)), this, SLOT(renameTab(int)));


    /* Button to add new tab. */
    m_loadXlsxFileButton = new QToolButton(m_xlsxTab);
    m_loadXlsxFileButton->setText(tr("Load xlsx file"));
    m_loadXlsxFileButton->setToolTip(tr("Select an xlsx file from your computer.\n"
                                        "All sheets will be loaded in this viewer,\n"
                                        "in addition to those already loaded.\n"
                                        "The last sheet loaded will be active."));
    connect(m_loadXlsxFileButton, SIGNAL(clicked()), this, SLOT(loadSheet()));
    m_xlsxTab->addTab(new QWidget(m_xlsxTab), tr(""));
    m_xlsxTab->setTabEnabled(0, false);
    m_xlsxTab->tabBar()->setTabButton(0, QTabBar::RightSide, m_loadXlsxFileButton);

    xlsxWidgetLayout->addWidget(m_xlsxTab);

    /* Set layout to main widget. */
    xlsxWidget->setLayout(xlsxWidgetLayout);

    /* Set main widget to dockwidget. */
    m_xlsxViewerDW->setWidget(xlsxWidget);

}

/* Create Toolbar */
void MainWindow::createToolBar(){

    m_toolBar = new QToolBar(this);
    m_toolBar->setMovable(false);

    /* Infobar */
    QLabel *infoLabel = new QLabel(m_toolBar);
    infoLabel->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::MinimumExpanding);
    infoLabel->setText(tr(APPLICATION_NAME) + tr(" - ")+ tr(APPLICATION_COMPANY) + tr(" - ") + tr(APPLICATION_YEAR));

    /* Help Button. */
    //QAction *help = new QAction(tr("Help"), m_toolBar);

    /* About Button. */
    QAction *about = new QAction(tr("About"), m_toolBar);
    connect(about, SIGNAL(triggered()), this, SLOT(about()));

    /* About Qt Button. */
    QAction *qtInfo = new QAction(tr("Qt"), m_toolBar);
    connect(qtInfo, SIGNAL(triggered()), qApp, SLOT(aboutQt()));

    /* Setup layout. */
    m_toolBar->addWidget(infoLabel);
    m_toolBar->addSeparator();
    //m_toolBar->addAction(help);
    //m_toolBar->addSeparator();
    m_toolBar->addAction(about);
    m_toolBar->addSeparator();
    m_toolBar->addAction(qtInfo);
    m_toolBar->addSeparator();

}

/*
 * [2] General methods.
 */

/* Return textversion of mail header. */
QString MainWindow::getMailHeader(int offset){

    QString txt;
    QStringList bcc_addresses = m_emailBcc->text().split(";");

    /* From/To. */
    txt += tr("From: ") + m_senderName->text() + tr(" <") + m_senderEmail->text() +  tr(">\n");
    txt += tr("To: <") + getData(m_emailColumnSelect->currentText(), offset) + m_emailAppendText->text() + tr(">\n");

    /* BCC. */
    foreach(QString bcc, bcc_addresses){
        if(!bcc.isEmpty()){
            txt += tr("Bcc: <") + bcc + tr(">\n");
        }
    }

    /* Subject. */
    txt += tr("Subject: [") + m_courseCode->text() + tr("] ") + m_emailSubject->text() + tr("\n");

    /* Global attachments. */
    for(int i = 0; i < m_attachments->count(); i++){
        QString info = tr("");
        QString filePath = m_attachments->itemData(i).toString();
        QFileInfo fInfo = QFileInfo(filePath);

        if(!fInfo.exists()){
            info += tr(" [!INVALID FILE]");
        }
        else{
            info += tr(" [") + QString::number(fInfo.size()/1024) + tr(" kB]");
        }

        txt += tr("Global Attachment: ") + m_attachments->itemText(i) + info + tr("\n");
    }

    /* Individual attachment. */
    if(m_attachmentColSelect->currentText() != tr("<none>")){
        QString info = tr("");
        QString fileName = getData(m_attachmentColSelect->currentText(), offset) + m_attachmentAppend->text();
        QString filePath = m_attachmentDirectory + QDir::separator() + fileName;
        QFileInfo fInfo = QFileInfo(filePath);

        if(!fInfo.exists()){
            info += tr(" [!INVALID FILE]");
        }
        else{
            info += tr(" [") + QString::number(fInfo.size()/1024) + tr(" kB]");
        }

        txt += tr("Individual Attachment: ") + fileName + info + tr("\n");
    }
    txt += tr("\n");

    return txt;

}

/* Parses the text from the editor to fill with values from the spreadsheet. */
QString MainWindow::getMailText(int offset){

    /* Get text from editor */
    QString txt = tr("");
    QTextEdit *te = qobject_cast<QTextEdit*>(m_textTab->currentWidget());
    if(te != NULL){
        txt = te->toPlainText();
    }

    /* Regexp to find all alphanumeric values between ## */
    QRegExp re("#([A-Z,a-z,0-9]*)#");

    /* Result text. */
    QString res;

    /* Match all values to be extracted from spreadsheet. */
    int lastpos = 0;
    int pos = 0;
    while((pos = re.indexIn(txt, pos)) != -1){
        /* Add part between last match and new match as-is. */
        res.append(txt.mid(lastpos, pos-lastpos));

        /* extract the matched text. */
        QString cell = re.cap(1);

        /* Get data. */
        res.append(getData(cell, offset));

        /* Keep positions between matches. */
        pos += re.matchedLength();
        lastpos = pos;
    }

    /* Add the remaining part of the text. */
    res.append(txt.mid(lastpos, txt.length()-lastpos));

    /* Return the parsed text. */
    return res;
}

/* Parse cell and get data from spreadsheet. */
QString MainWindow::getData(QString cell, int offset)
{
    /* Extract column and row. */
    QRegExp regex("([A-Z,a-z]+)([0-9]*)");

    if(regex.indexIn(cell, 0) != -1){

        /* Column. Convert to uppercase. */
        QString colText = regex.cap(1).toUpper();

        /* Row. */
        QString rowText = regex.cap(2);

        /* Rows can be static or dynamic. */
        int row = 0;
        if(rowText.isEmpty()){
            row = offset;
        }
        else{
            row = rowText.toInt();
        }

        /*
         * Columns are [a-z] and should be parsed to an integer.
         *
         * The integer value is the ascii value minus the ascii value
         * of character 'A' + 1. The position adds a power of 26.
         *
         * Currently we can parse columns up to length 4.
         *
         * TODO: find a more generic algorithm.
         */
        int col = 0;
        switch(colText.length()){
          case 4:
            col = (qPow(26, 3) * (colText.data()[0].toLatin1() - 'A' + 1)) +
                  (qPow(26, 2) * (colText.data()[1].toLatin1() - 'A' + 1)) +
                  (qPow(26, 1) * (colText.data()[2].toLatin1() - 'A' + 1)) +
                  (qPow(26, 0) * (colText.data()[3].toLatin1() - 'A' + 1));
            break;
          case 3:
            col = (qPow(26, 2) * (colText.data()[0].toLatin1() - 'A' + 1)) +
                  (qPow(26, 1) * (colText.data()[1].toLatin1() - 'A' + 1)) +
                  (qPow(26, 0) * (colText.data()[2].toLatin1() - 'A' + 1));
            break;
          case 2:
            col = (qPow(26, 1) * (colText.data()[0].toLatin1() - 'A' + 1)) +
                  (qPow(26, 0) * (colText.data()[1].toLatin1() - 'A' + 1));
            break;
          case 1:
            col = colText.data()[0].toLatin1() - 'A' + 1;
            break;
          default:
            return QString("[INV_REF!]");
        }

        /* Get the data at row, col. */
        return getData(row, col);
    }

    /* Return invalid if we cannot parse this cell properly. */
    return QString("[INV_REF!]");
}

/* Extracts the data from the spreadsheet at row,col */
QString MainWindow::getData(int row, int col){

    /* Get a pointer to the data. */
    QTableView *d = (QTableView*)m_xlsxTab->currentWidget();

    /*
     * Check row and col, if we have a proper spreadsheet and
     * if the value is in range. Return invalid if not.
     */
    if(row < 0 || col < 0 || d == NULL || m_xlsxTab->currentIndex() == 0 || d->model() == NULL || col > d->model()->columnCount() || row > d->model()->rowCount()){
        return QString("[INV_REF!]");
    }

    /* Return the value at row,col */
    return(d->model()->data(d->model()->index(row-1, col-1)).toString());
}

/* Validate email address. */
bool MainWindow::isValidEmail(QString address){
    return QRegExp("[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63}", Qt::CaseInsensitive).exactMatch(address);
}

/* Validate HR student email address. */
bool MainWindow::isValidHRStudentEmail(QString address){
    return QRegExp("\\d{7}@hr.nl", Qt::CaseInsensitive).exactMatch(address);
}

/* Validate HR employee email address. */
bool MainWindow::isValidHREmployeeEmail(QString address){
    return QRegExp("[a-z]{5}@hr.nl", Qt::CaseInsensitive).exactMatch(address);
}

/*
 * [3] Save/load/delete Settings.
 */

/* Save settings. */
void MainWindow::saveSettings(){

    /* Get settings. */
    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR), this);

    /* Settings parameters. */
    s->setValue(tr("saveOnExit"), m_saveOnExitCheckBox->isChecked());
    s->setValue(tr("validateHR"), m_validateHR->isChecked());
    s->setValue(tr("runtimeValidate"), m_runtimeValidate->isChecked());

    /* Email parameters. */
    s->setValue(tr("senderName"), m_senderName->text());
    s->setValue(tr("senderEmail"), m_senderEmail->text());
    s->setValue(tr("emailBcc"), m_emailBcc->text());
    s->setValue(tr("emailSubject"), m_emailSubject->text());
    s->setValue(tr("emailAppend"), m_emailAppendText->text());
    s->setValue(tr("courseCode"), m_courseCode->text());
    s->setValue(tr("SMTPserver"), m_SMTPserver->text());
    s->setValue(tr("SMTPport"), m_SMTPport->text());
    s->setValue(tr("SMTPtype"), m_SMTPtype->currentText());

    /* Texts. */
    s->beginWriteArray(tr("mailTexts"));
    int index = 0;
    for(int i = 0; i < m_textTab->count(); i++){
        if(QTextEdit *te = qobject_cast<QTextEdit*>(m_textTab->widget(i))){
            if(!te->toPlainText().isEmpty()){
                s->setArrayIndex(index);
                s->setValue(tr("name"), m_textTab->tabText(i));
                s->setValue(tr("text"), te->toPlainText());
                index++;
            }
        }
    }
    s->endArray();

    /* Save. */
    s->sync();
}

/* Load settings. */
void MainWindow::loadSettings(){

    /* Get settings. */
    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR), this);

    /* Settings parameters. */
    m_saveOnExitCheckBox->setChecked(s->value(tr("saveOnExit"), QVariant(false)).toBool());
    m_validateHR->setChecked(s->value(tr("validateHR"), QVariant(true)).toBool());
    m_runtimeValidate->setChecked(s->value(tr("runtimeValidate"), QVariant(true)).toBool());

    /* Email parameters. */
    m_senderName->setText(s->value(tr("senderName"), tr("")).toString());
    m_senderEmail->setText(s->value(tr("senderEmail"), tr("")).toString());
    m_emailBcc->setText(s->value(tr("emailBcc"), tr("")).toString());
    m_emailSubject->setText(s->value(tr("emailSubject"), tr("")).toString());
    m_emailAppendText->setText(s->value(tr("emailAppend"), tr("@hr.nl")).toString());
    m_courseCode->setText(s->value(tr("courseCode"), tr("")).toString());
    m_SMTPserver->setText(s->value(tr("SMTPserver"), tr("smtp.hr.nl")).toString());
    m_SMTPport->setText(s->value(tr("SMTPport"), tr("465")).toString());
    m_SMTPtype->setCurrentText(s->value(tr("SMTPtype"), tr("SSL")).toString());

    /* Texts. */
    int num = s->beginReadArray(tr("mailTexts"));
    for(int i = 0; i < num; i++){
        s->setArrayIndex(i);
        QString name = s->value(tr("name"), tr("unknown")).toString();
        addNewTextTab();
        m_textTab->setTabText(m_textTab->currentIndex(), name);
        ((QTextEdit*)m_textTab->currentWidget())->setText(s->value(tr("text"), tr("")).toString());
    }
    if(num == 0){
        addNewTextTab();
    }
    s->endArray();

}

/* Delete saved settings to get a fresh application. */
void MainWindow::deleteSettings(){

    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR), this);

    /* Sure? */
    if(QMessageBox::question(this, tr("Delete all settings?"),
                                   tr("Are you sure you want to delete all items and saved settings?\n"
                                      "")
                             ) != QMessageBox::Yes){
        return;
    }

    /* Clear saved settings. */
    s->clear();
    s->sync();

    /* Remove and delete all texts. */
    int num = m_textTab->count() - 1;
    for(int i = num; i >= 0; i--){
        QTextEdit *te;
        if((te = qobject_cast<QTextEdit*>(m_textTab->widget(i)))){
            m_textTab->removeTab(i);
            delete te;
        }
    }

    /* Remove and delete all spreadsheet tabs. */
    num = m_xlsxTab->count() - 1;
    for(int i = num; i >= 0; i--){
        QTableView *tv;
        if((tv = qobject_cast<QTableView*>(m_xlsxTab->widget(i)))){
            m_xlsxTab->removeTab(i);
            delete tv;
        }
    }

    /* Set default values. */
    loadSettings();

}

/* When the application closes we might want to save the settings. */
void MainWindow::closeEvent(QCloseEvent *closeEvent){

    /* Should we save settings? */
    if(m_saveOnExitCheckBox->isChecked()){
        QMessageBox::StandardButton answer = QMessageBox::question(this,
                                                                   tr("Save settings?"),
                                                                   tr("Do you want to save your email texts and general parameters?\n"),
                                                                   QMessageBox::Cancel | QMessageBox::No | QMessageBox::Yes,
                                                                   QMessageBox::Cancel);
        /* Do not close. */
        if(answer == QMessageBox::Cancel){
            closeEvent->ignore();
            return;
        }

        /* Close without saving. */
        if(answer == QMessageBox::No){
            closeEvent->accept();
            return;
        }
    }

    /* Save */
    saveSettings();

    /* Close */
    closeEvent->accept();

}

/*
 * [4] Show/hide widgets.
 */

/* Show/hide the settingsWidget. */
void MainWindow::toggleSettingsWidget(bool s){

    if(!s){
        /* Hide */
        m_toggleSettingsAnimation->setStartValue(400);
        m_toggleSettingsAnimation->setEndValue(0);
        m_settingsWidgetToggleButton->setText(tr(">"));
        m_settingsWidgetToggleButton->setToolTip(tr("Show settings"));
    }
    else{
        /* Show. */
        m_toggleSettingsAnimation->setStartValue(0);
        m_toggleSettingsAnimation->setEndValue(400);
        m_settingsWidgetToggleButton->setText(tr("<"));
        m_settingsWidgetToggleButton->setToolTip(tr("Hide settings"));
    }

    m_toggleSettingsAnimation->start();

}

/* Show/hide the SMTPWidget. */
void MainWindow::toggleSMTPWidget(bool s){

    if(!s){
        /* Hide */
        m_toggleSMTPAnimation->setStartValue(400);
        m_toggleSMTPAnimation->setEndValue(0);
        m_SMTPWidgetToggleButton->setText(tr("<"));
        m_SMTPWidgetToggleButton->setToolTip(tr("Show SMTP options"));
    }
    else{
        /* Show. */
        m_toggleSMTPAnimation->setStartValue(0);
        m_toggleSMTPAnimation->setEndValue(400);
        m_SMTPWidgetToggleButton->setText(tr(">"));
        m_SMTPWidgetToggleButton->setToolTip(tr("Hide SMTP options"));
    }

    m_toggleSMTPAnimation->start();

}

/* Show/hide the generateWidget. */
void MainWindow::toggleGenerateWidget(bool s){

    if(!s){
        /* Hide. */
        m_toggleGenerateAnimation->setStartValue(400);
        m_toggleGenerateAnimation->setEndValue(0);
        m_generateWidgetToggleButton->setText(tr("<"));
        m_generateWidgetToggleButton->setToolTip(tr("Show options"));
    }
    else{
        /* Show. */
        m_toggleGenerateAnimation->setStartValue(0);
        m_toggleGenerateAnimation->setEndValue(400);
        m_generateWidgetToggleButton->setText(tr(">"));
        m_generateWidgetToggleButton->setToolTip(tr("Hide options"));
    }

    /* Start animation. */
    m_toggleGenerateAnimation->start();

}

/* Show/hide the generateWidget. */
void MainWindow::toggleAttachmentWidget(bool s){

    if(!s){
        /* Hide. */
        m_toggleAttachmentAnimation->setStartValue(100);
        m_toggleAttachmentAnimation->setEndValue(0);
        m_attachmentWidgetToggleButton->setText(tr("^"));
        m_attachmentWidgetToggleButton->setToolTip(tr("Show individual attachment options."));
    }
    else{
        /* Show. */
        m_toggleAttachmentAnimation->setStartValue(0);
        m_toggleAttachmentAnimation->setEndValue(100);
        m_attachmentWidgetToggleButton->setText(tr("v"));
        m_attachmentWidgetToggleButton->setToolTip(tr("Hide individual attachment options."));
    }

    /* Start animation. */
    m_toggleAttachmentAnimation->start();

}

/*
 *  [5] Update.
 */

/* Add attachment file dialog. */
void MainWindow::addAttachment(){

    /* Ask for the file. */
    QString filePath = QFileDialog::getOpenFileName(this, "Select file");
    if(filePath.isEmpty()){
        return;
    }

    /* Get info. */
    QFileInfo fInfo = QFileInfo(filePath);

    /* Add filename to the list of attachments. */
    m_attachments->addItem(fInfo.fileName(), filePath);
    m_attachments->setItemData(m_attachments->count() - 1, tr("Size: ") + QString::number(fInfo.size()/1024) + tr("kB.\n\nFull path:\n") + filePath + tr("\n"), Qt::ToolTipRole);

    /* Load values into preview text. */
    updateText();

    /* Show attachment fields. */
    m_attachments->show();
    m_deleteSelectedAttachment->show();

}

/* Delete attachment from list. */
void MainWindow::deleteAttachment(){

    /* Is there something to delete? */
    if(m_attachments->currentText().isEmpty()){
        return;
    }

    /* Sure to delete? */
    if(QMessageBox::question(this, tr("Remove Attachment?"),
                                   tr("Are you sure you want to remove the attachment \"") +
                                   m_attachments->currentText() + tr("\"?\n") +
                                   tr("\nFull path:\n") + m_attachments->currentData().toString() + tr("\n")
                             ) != QMessageBox::Yes){
        return;
    }

    /* Ok, then delete it. */
    m_attachments->removeItem(m_attachments->currentIndex());

    updateText();

    /* Hide fields when there are no attachments. */
    if(m_attachments->count() < 1){
        m_attachments->hide();
        m_deleteSelectedAttachment->hide();
    }

}

void MainWindow::selectAttachmentDirectory(){

    m_attachmentDirectory = QFileDialog::getExistingDirectory(this, tr("Select Directory to load attachments from:"));
    m_selectAttachmentDirectoryButton->setToolTip(tr("Select the directory where the\n"
                                                     "individual attachments are located.\n\n"
                                                     "Current directory:\n") + m_attachmentDirectory);

    updateText();

}

/* Load an xlsx sheet */
void MainWindow::loadSheet(){

    /* Ask for the file. */
    QString filePath = QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), "*.xlsx");
    if(filePath.isEmpty()){
        return;
    }

    /* Open the document. */
    QXlsx::Document *xlsx = new QXlsx::Document(filePath, m_xlsxTab);

    /* Add add sheets (tabs) to the viewer. */
    foreach(QString sheetName, xlsx->sheetNames()){

        /* Load sheet */
        QXlsx::Worksheet *sheet = dynamic_cast<QXlsx::Worksheet *>(xlsx->sheet(sheetName));
        if(sheet){

            /* Create a tableview for this sheet. */
            QTableView *view = new QTableView(m_xlsxTab);
            view->setToolTip(tr("This is the data (read-only) from the selected sheet\n"
                                "that will be used to generate the e-mail from."));

            /* Set to read-only. */
            view->setEditTriggers(QAbstractItemView::NoEditTriggers);
            view->setModel(new QXlsx::SheetModel(sheet, view));

            /* Handle merged cells. */
            foreach (QXlsx::CellRange range, sheet->mergedCells()){
                view->setSpan(range.firstRow()-1, range.firstColumn()-1, range.rowCount(), range.columnCount());
            }

            /* Add sheet as a tab to viewer. */
            int tabIndex = m_xlsxTab->addTab(view, sheetName);
            m_xlsxTab->setCurrentIndex(tabIndex);
        }
    }

}

/* Slot called when selecting an onther sheet. */
void MainWindow::updateSheet(){

    /* Get a pointer to the data. */
    QTableView *d = (QTableView*)m_xlsxTab->currentWidget();

    /* Extract columns and rows. */
    QStringList columnNames;
    QStringList rowNames;
    if(d != NULL && m_xlsxTab->currentIndex() != 0){
        for(int i = 0; i < d->model()->columnCount(); i++){
            columnNames.append(d->model()->headerData(i, Qt::Horizontal).toString());
        }
        for(int i = 1; i <= d->model()->rowCount(); i++){
            rowNames.append(QString::number(i));
        }
    }

    /* Insert extracted values into selection boxes. */
    QString tmp = m_emailColumnSelect->currentText();
    m_emailColumnSelect->clear();
    m_emailColumnSelect->addItems(columnNames);
    m_emailColumnSelect->setCurrentText(tmp);

    /* Now do those who may be ignored. */
    columnNames.prepend(tr("<none>"));
    rowNames.prepend(tr("<none>"));

    QComboBox *cols[] = {m_nameColSelect, m_finalGradeColSelect, m_startColSelect, m_stopColSelect, m_attachmentColSelect};
    for(int i = 0; i < 5; i++){
        QComboBox *c = cols[i];
        QString tmp = c->currentText();
        c->clear();
        c->addItems(columnNames);
        c->setCurrentText(tmp);
    }

    QComboBox *rows[] = {m_headerRowSelect, m_maxRowSelect};
    for(int i = 0; i < 2; i++){
        QComboBox *c = rows[i];
        QString tmp = c->currentText();
        c->clear();
        c->addItems(rowNames);
        c->setCurrentText(tmp);
    }

    /* Reload values. */
    updateInfo();

}

/* Update dependend items when values change. */
void MainWindow::updateInfo(){
    /* Block (circular) updates. */
    blockRowSignals(true);

    /* Get pointer to data and find values. */
    QTableView *d = (QTableView*)m_xlsxTab->currentWidget();
    int max = 0;
    int start = 1;
    int stop = max;
    int preview = 1;

    if(d != NULL && m_xlsxTab->currentIndex() != 0){
        max = d->model()->rowCount();
        stop = max;
    }

    if(!m_firstRowSelect->currentText().isEmpty()){
        start = m_firstRowSelect->currentText().toInt();
    }
    if(!m_lastRowSelect->currentText().isEmpty()){
        stop = m_lastRowSelect->currentText().toInt();
    }
    if(!m_previewSelect->currentText().isEmpty()){
        preview = m_previewSelect->currentText().toInt();
    }

    /* Remove items from selection boxes. */
    m_firstRowSelect->clear();
    m_lastRowSelect->clear();
    m_previewSelect->clear();

    /* Insert new items to selection boxes. */
    for(int i = 1; i <= max; i++){
        m_firstRowSelect->addItem(QString::number(i));
        if(i >= start && i <= stop && !getData(m_emailColumnSelect->currentText(), i).isEmpty()){
            m_previewSelect->addItem(QString::number(i));
        }
        if(i >= start){
            m_lastRowSelect->addItem(QString::number(i));
        }
    }

    /* Set old values if useful. */
    m_firstRowSelect->setCurrentText(QString::number(start));
    if(stop >= start){
        m_lastRowSelect->setCurrentText(QString::number(stop));
    }
    else{
        m_lastRowSelect->setCurrentText(QString::number(max));
    }
    m_previewSelect->setCurrentText(QString::number(preview));

    /* Re-enable updates. */
    blockRowSignals(false);

    /* Calculate and set number of generated mails. */
    if(max > 0){
        m_nMailsDisplay->display(m_previewSelect->count());
    }
    else{
        m_nMailsDisplay->display(0);
    }

    /* Update preview text. */
    updateText();

}

/* Update and parse values. */
void MainWindow::updateText(){

    QStringList bcc_addresses = m_emailBcc->text().split(";");
    QStringList report_addresses = m_reportCC->text().split(";");

    /* Generate preview text. */
    QString res;

    /* Get the mail number. */
    int offset = m_previewSelect->currentText().toInt();

    if(m_runtimeValidate->isChecked()){
        /* Check option fields and make red when useless. */
        !isValidEmail(m_senderEmail->text()) ? m_senderEmail->setStyleSheet(tr("background-color: #FF9999;")) :
                                               m_senderEmail->setStyleSheet(tr(""));

        if(m_validateHR->isChecked()){
            !isValidHREmployeeEmail(m_senderEmail->text()) ? m_senderEmail->setStyleSheet(tr("background-color: #FF9999;")) :
                                                            m_senderEmail->setStyleSheet(tr(""));
        }

        m_emailSubject->text().length() < 3  ? m_emailSubject->setStyleSheet(tr("background-color: #FF9999;")) :
                                               m_emailSubject->setStyleSheet(tr(""));

        m_courseCode->text().length() < 3    ? m_courseCode->setStyleSheet(tr("background-color: #FF9999;")) :
                                               m_courseCode->setStyleSheet(tr(""));

        m_xlsxTab->count() < 2               ? m_loadXlsxFileButton->setStyleSheet(tr("background-color: #FF9999;")) :
                                               m_loadXlsxFileButton->setStyleSheet(tr(""));

        foreach(QString bcc, bcc_addresses){
            !bcc.isEmpty() && !isValidEmail(bcc) ? m_emailBcc->setStyleSheet(tr("background-color: #FF9999;")) :
                                                   m_emailBcc->setStyleSheet(tr(""));
        }
        foreach(QString reportcc, report_addresses){
            !reportcc.isEmpty() && !isValidEmail(reportcc) ? m_reportCC->setStyleSheet(tr("background-color: #FF9999;")) :
                                                             m_reportCC->setStyleSheet(tr(""));
        }
    }
    else{
        m_senderEmail->setStyleSheet(tr(""));
        m_emailSubject->setStyleSheet(tr(""));
        m_courseCode->setStyleSheet(tr(""));
        m_loadXlsxFileButton->setStyleSheet(tr(""));
        m_emailBcc->setStyleSheet(tr(""));
        m_reportCC->setStyleSheet(tr(""));
    }

    /* Get the main contents. */
    res += getMailHeader(offset);
    res += getMailText(offset);

    /* set text to preview. */
    m_previewText->setText(res);

}

/* Hacky thing to avoid circular updates when loading new values into the selection boxes. */
void MainWindow::blockRowSignals(bool b){

    m_firstRowSelect->blockSignals(b);
    m_lastRowSelect->blockSignals(b);
    m_previewSelect->blockSignals(b);

}

/*
 * [6] Generate mail contents.
 */

/* Add new composer tab and make it active. */
void MainWindow::addNewTextTab(){

    QTextEdit *newText = new QTextEdit(this);
    newText->setToolTip(tr("In this box you can type any text you want as a\n"
                           "template for the generated emails.\n\n"
                           "Text between ## will be parsed from the selected\n"
                           "spreadsheet as follows:\n"
                           " #A# \twill include the value in column A. It will be\n"
                           "     \tdifferent for all emails.\n"
                           " #A1#\twill include the cell A1 in the emails. It will\n"
                           "     \tbe the same in all emails.\n\n"
                           "Note:\n"
                           "Only values will be read from the spreadsheet, not the formatting,\n"
                           "so if you want to use rounded values use the ROUND() function\n"
                           "before loading the spreadsheet."));
    connect(newText, SIGNAL(textChanged()), this, SLOT(updateText()));

    int num = m_textTab->count() - 1;
    int newNum = 1;
    for(int i = num; i >= 0; i--){
        QTextEdit *te;
        if((te = qobject_cast<QTextEdit*>(m_textTab->widget(i)))){
            QRegExp re("[T](\\d+)");
            QString name = m_textTab->tabText(i);
            if(re.exactMatch(name)){
                int oldNum = re.cap(1).toInt();
                if(oldNum >= newNum){
                    newNum = oldNum + 1;
                }
            }
        }
    }

    m_textTab->addTab(newText, tr("T") + QString::number(newNum));
    m_textTab->setCurrentWidget(newText);

}

/* Generate template based on selected rows and columns. */
void MainWindow::generateText(bool newTab){

    /* Should we include names and maximum points / default values? */
    bool header = !(m_headerRowSelect->currentText().startsWith(tr("<")));
    bool maxpoints = !(m_maxRowSelect->currentText().startsWith(tr("<")));

    /* Start generating text. */
    QString txt = tr("Beste ");

    /* Add name or default value? */
    if(m_nameColSelect->currentText().startsWith(tr("<"))){
        txt += tr("Student,\n\n");
    }
    else{
        txt += tr("#") + m_nameColSelect->currentText() + tr("#,\n\n");
    }

    /* Add final grade? */
    if(!(m_finalGradeColSelect->currentText().startsWith(tr("<")))){

        txt += tr("Je cijfer ");

        /* Include coursecode? */
        if(!m_courseCode->text().isEmpty()){
            txt += tr("voor het vak ") + m_courseCode->text() + tr(" ");
        }

        txt += tr("is een #") + m_finalGradeColSelect->currentText() + tr("#");

        txt += tr(".\n\n");
    }

    /* Add columns with values. */
    if(!(m_startColSelect->currentText().startsWith(tr("<"))) && !(m_stopColSelect->currentText().startsWith(tr("<"))) && m_startColSelect->currentIndex() <= m_stopColSelect->currentIndex()){

        /* Do we have a header/name? */
        if(maxpoints){
            txt += tr("Specificatie (punten/totaal):\n");
        }
        else{
            txt += tr("Specificatie:\n");
        }

        /* Add all columns. */
        for(int i = m_startColSelect->currentIndex(); i <= m_stopColSelect->currentIndex(); i++){

            /* Include header? */
            if(header){
                txt += tr("#") + m_startColSelect->itemText(i) + m_headerRowSelect->currentText() + tr("#: ");
            }

            txt += tr("#") + m_startColSelect->itemText(i) + tr("#");

            /* Include maximum points/default value? */
            if(maxpoints){
                txt += tr("/#") + m_startColSelect->itemText(i) + m_maxRowSelect->currentText() + tr("#");
            }

            txt += tr("\n");
        }
    }

    /* Add greeting. */
    txt += tr("\n");
    txt += tr("Met vriendelijke groet,\n\n") + m_senderName->text() + tr("\n");

    /* Disclaimer. */
    txt += tr("\n--\n"
              "This message was automatically generated and is\n"
              "intended to be informative only. No rights can be\n"
              "claimed based on the contents of this message."
              "\n--\n");

    /* Create new tab? */
    if(newTab){
        addNewTextTab();
    }

    /* Set text. */
    ((QTextEdit*)m_textTab->currentWidget())->setText(txt);

    /* We're done! Hide the generate widget. */
    m_generateWidgetToggleButton->setChecked(false);

}

/*
 * [7] Close or rename a tab in the XLSX viewer or the Editor.
 */

/* Close tab in a tabwidget */
void MainWindow::closeTab(int index){

    /* Get the tabwidget */
    QTabWidget *tw = (QTabWidget*)sender();

    /* Get the tab. */
    QWidget *tab = tw->widget(index);

    /* The editor cannot have 0 tabs. */
    if(tw == m_textTab && tw->count() < 3){
        return;
    }

    /* Sure? */
    if(QMessageBox::question(this, tr("Close tab?"),
                                   tr("Are you sure you want to close the tab \"") +
                                   tw->tabText(index) +
                                   tr("\"?")
                             ) != QMessageBox::Yes){
        return;
    }

    /* Remove and delete tab. */
    tw->removeTab(index);
    delete tab;

}

/* Rename a tab in the editor. */
void MainWindow::renameTab(int index){

    /* Get the calling tabWidget. */
    QTabWidget *tw = (QTabWidget*)this->sender();

    if(index == 0){
        return;
    }

    /* Ask for new name. */
    bool ok;
    QString oldname = tw->tabText(index);
    QString name = QInputDialog::getText(this, tr("Set new name for tab:"), tr("New name:"), QLineEdit::Normal, oldname, &ok);

    /* If ok, then rename. */
    if(ok && !name.isEmpty()){
        tw->setTabText(index, name);
    }
}

/*
 * [8] SMTP.
 */

/* Connect to SMTP server. */
void MainWindow::SMTPconnect(){
    bool ok;

    /* Get servername. */
    QString smtpServer = m_SMTPserver->text();

    /* Get portnumber. */
    int smtpPort = m_SMTPport->text().toInt();

    /* Get connection type. */
    SmtpClient::ConnectionType smtpType = static_cast<SmtpClient::ConnectionType>(m_SMTPtype->currentData().toInt());

    /* Lib likes to throw exceptions... */
    try {
        /* Try to set values. */
        m_SMTPConnection = new SmtpClient(smtpServer, smtpPort, smtpType);
    }
    catch (...){

    }

    /*
     * Get login name for SMTP server.
     *
     * Default is the sender email address.
     */
    QString user = QInputDialog::getText(this, tr("Username:"),
                                           tr("SMTP username for ") + smtpServer +
                                           tr(":") + QString::number(smtpPort),
                                           QLineEdit::Normal, m_senderEmail->text(),
                                         &ok);

    /* Value not empty and user dit not press close or cancel? */
    if(!ok || user.isEmpty()){
        SMTPdisconnect();
        return;
    }

    /* Ask for SMTP password. */
    QString password = QInputDialog::getText(this, tr("Password:"),
                                               tr("SMTP Password for <") + user +
                                               tr(">@") + smtpServer +
                                               tr(":") + QString::number(smtpPort),
                                               QLineEdit::Password, tr(""),
                                             &ok);

    /* Value not empty and user dit not press close or cancel? */
    if(!ok || password.isEmpty()){
        SMTPdisconnect();
        return;
    }

    /* Set username and password. */
    m_SMTPConnection->setUser(user);
    m_SMTPConnection->setPassword(password);

    /* Lib throws exceptions... */
    try {
        /* Connect to SMTP server. */
        if(!m_SMTPConnection->connectToHost()){
            QMessageBox::warning(this, tr("SMTP Connect"), tr("Could not connect to SMTP server!"));
            SMTPdisconnect();
            return;
        }

        /* Login. */
        if(!m_SMTPConnection->login()){
            QMessageBox::warning(this, tr("SMTP Connect"), tr("SMTP login failed! Wrong username/password."));
            SMTPdisconnect();
            return;
        }
    }
    catch (...){

    }

}

/* Disconnect. */
void MainWindow::SMTPdisconnect(){
    if(m_SMTPConnection != NULL){

        /* Should include this, but throws uncatchable exceptions. */
        //m_smtp->quit();

        delete m_SMTPConnection;
    }

    /* Set to NULL for next connect. */
    m_SMTPConnection = NULL;
}

/*
 * The main thing.. Sending emails.
 *
 * TODO: A bit long and not very well structured.
 *
 */
void MainWindow::sendMails(){

    /* Calculate number of mails. */
    int nMails = m_previewSelect->count();
    int nAttachments = 0;

    /* Display Progress. */
    QWidget progress(this);
    progress.setFixedSize(this->width(), this->height());
    progress.setAutoFillBackground(true);

    QVBoxLayout progressLayout(&progress);
    progressLayout.setAlignment(Qt::AlignHCenter);

    QLabel progressText(this);
    progressText.setFixedSize(this->width(), this->height());
    progressText.setAutoFillBackground(true);
    progressText.setAlignment(Qt::AlignHCenter | Qt::AlignTop);

    QProgressBar progressBar(this);
    progressBar.setRange(0, 0);
    progressBar.setFixedWidth((this->width()*2)/3);

    progressLayout.addSpacing(this->height()/8);
    progressLayout.addWidget(&progressBar);
    progressLayout.addWidget(&progressText);
    progressLayout.setAlignment(&progressBar, Qt::AlignHCenter);
    progress.setLayout(&progressLayout);
    progress.show();

    /* Mail Objects */
    MimeMessage messages[nMails];
    MimeText texts[nMails];
    EmailAddress sender(tr(""));

    /* Check parameters... */
    progressText.setText(tr("Checking parameters..."));
    qApp->processEvents();

    /* Check sender. */
    QString fromEmail = m_senderEmail->text();
    if(!isValidEmail(fromEmail)){
        QMessageBox::warning(this, tr("Error:"), tr("Sender email address is invalid!"));
        m_senderEmail->setFocus();
        return;
    }
    sender.setAddress(fromEmail);

    /* Use name or email address as sender. */
    QString fromName = m_senderName->text();
    if(fromName.isEmpty()){
        fromName = fromEmail;
    }
    sender.setName(fromName);

    /* Check course code. */
    QString coursecode = m_courseCode->text();
    if(coursecode.length() < 2){
        QMessageBox::warning(this, tr("Error:"), tr("Course code cannot be less than 2 characters!"));
        m_courseCode->setFocus();
        return;
    }

    /* Check subject. */
    QString subject = tr("[") + coursecode + tr("] ") + m_emailSubject->text();
    if(m_emailSubject->text().length() < 2){
        QMessageBox::warning(this, tr("Error:"), tr("Subject cannot be less than 2 characters!"));
        m_emailSubject->setFocus();
        return;
    }

    /* Check if there are any messages to send. */
    if(m_nMailsDisplay->value() == 0){
        QMessageBox::warning(this, tr("Error:"), tr("The number of messages is 0!"));
        return;
    }

    /* Get bcc addresses. */
    QStringList bcc_addresses = m_emailBcc->text().split(";");
    /* Check if the bcc's are OK. */
    foreach(QString bcc, bcc_addresses){
        if(bcc.isEmpty()){
            bcc_addresses.removeAll(bcc);
            continue;
        }

        if(!isValidEmail(bcc)){
            QMessageBox::warning(this, tr("Error:"), tr("The bcc email address ") + bcc + tr(" is invalid!"));
            m_emailBcc->setFocus();
            return;
        }
    }

    /* Get report cc addresses. */
    QStringList report_cc_addresses = m_reportCC->text().split(";");
    /* Check if the cc's are OK. */
    foreach(QString cc, report_cc_addresses){
        if(cc.isEmpty()){
            report_cc_addresses.removeAll(cc);
            continue;
        }

        if(!isValidEmail(cc)){
            QMessageBox::warning(this, tr("Error:"), tr("The Report CC email address ") + cc + tr(" is invalid!"));
            m_reportCC->setFocus();
            return;
        }
    }

    /* Read Attachments. */
    QList<QFile*> attachmentFiles;
    QList<MimeAttachment*> attachments;
    for(int i = 0; i < m_attachments->count(); i++){
        QString fileName = m_attachments->itemData(i).toString();
        QFile *f = new QFile(fileName);
        if(f->exists()){
            MimeAttachment *att = new MimeAttachment(f);
            attachments.append(att);
            nAttachments++;
        }
        else{
            QMessageBox::warning(this, tr("Error:"), tr("Attachment ") + fileName + tr(" can not be loaded!"));
            return;
        }
    }
    if(m_attachmentColSelect->currentText() != tr("<none>")){
        nAttachments++;
    }


    /* Checking messages... */
    progressText.setText("Checking messages...");
    qApp->processEvents();

    /* Check and generate mails. */
    for(int i = 0; i < nMails; i++){
        int rowIndex = m_previewSelect->itemText(i).toInt();

        /* Set sender. */
        messages[i].setSender(&sender);

        /* Recipient address OK? */
        QString recv_mail = getData(m_emailColumnSelect->currentText(), rowIndex) + m_emailAppendText->text();
        if(!isValidEmail(recv_mail)){
            QMessageBox::warning(this, tr("Error:"), tr("The email address ") + recv_mail +
                                 tr(" on line ") + QString::number(rowIndex) + tr(" is invalid!"));
            m_emailColumnSelect->setFocus();
            return;
        }
        if(m_validateHR->isChecked()){
            if(!isValidHRStudentEmail(recv_mail)){
                QMessageBox::warning(this, tr("Error:"), tr("The email address ") + recv_mail +
                                     tr(" on line ") + QString::number(rowIndex) +
                                     tr(" is not a valid HR student email address!"));
                return;
            }
        }

        /* Add receiver. */
        messages[i].addTo(new EmailAddress(recv_mail));

        /* Mailtext OK? */
        QString mailText = getMailText(rowIndex);
        if(mailText.contains("[INV_REF!]")){
            QMessageBox::warning(this, tr("Error:"), tr("There are invalid references in the mailtext of email ") +
                                 QString::number(rowIndex) + tr("!"));
            return;
        }

        /* Add text to mail. */
        texts[i].setText(mailText);
        messages[i].addPart(&texts[i]);

        /* Add subject. */
        messages[i].setSubject(subject);

        /* Add bcc's */
        foreach(QString bcc, bcc_addresses){
            messages[i].addBcc(new EmailAddress(bcc));
        }

        /* Add attachments. */
        foreach(MimeAttachment *att, attachments){
            messages[i].addPart(att);
        }

        /* Add individual attachment. */
        if(m_attachmentColSelect->currentText() != tr("<none>")){
            QString fileName = m_attachmentDirectory + QDir::separator() + getData(m_attachmentColSelect->currentText(), rowIndex) + m_attachmentAppend->text();
            QFile *f = new QFile(fileName);
            if(f->exists()){
                MimeAttachment *att = new MimeAttachment(f);
                messages[i].addPart(att);
            }
            else{
                QMessageBox::warning(this, tr("Error:"), tr("Attachment ") + fileName + tr(" can not be loaded!"));
                return;
            }
        }
    }

    /* Connect to SMTP */
    progressText.setText("Connect to SMTP server...");
    qApp->processEvents();

    /* Do we already have a connection? If not, connect. */
    if(DO_NOT_SEND_EMAILS == 0 && m_SMTPConnection == NULL){
        SMTPconnect();
        if(m_SMTPConnection == NULL){
            return;
        }
    }

    /* Confirm mails. */
    progressText.setText("Confirm...");
    qApp->processEvents();

    /* Sure? */
    if(QMessageBox::question(this, tr("Send emails now?"),
                                   tr("Are you sure you want to send ") +
                                   QString::number(nMails) +
                                   tr(" emails with the subject: \"") + subject +
                                   tr("\" and ") + QString::number(nAttachments) + tr(" attachments now?")
                             ) != QMessageBox::Yes){
        return;
    }

    /* Set progressbar range. */
    progressBar.setRange(0, nMails);

    /* Statistics. */
    QString success;
    int nSuccess = 0;
    QString failed;
    int nFailed = 0;
    QString allTexts;

    /* Send messages. */
    for(int i = 0; i < nMails; i++){
        int rowIndex = m_previewSelect->itemText(i).toInt();

        progressText.setText(tr("Sending message ") + QString::number(i+1) + tr(" / ") + QString::number(nMails) + tr("..."));
        progressBar.setValue(i);
        qApp->processEvents();

        /* Add contents. */
        allTexts.append(tr("\n\n============================== " ) + QString::number(rowIndex) + tr(" ==============================\n"));
        allTexts.append(getMailHeader(rowIndex));
        allTexts.append(texts[i].getText());

        /* Try to send the mail. If failed, keep track of this. */
        if(!sendMail(&messages[i])){
            failed += tr("  ") + messages[i].getRecipients()[0]->getAddress() + tr("\n");
            nFailed++;
            continue;
        }

        success += tr("  ") + messages[i].getRecipients()[0]->getAddress() + tr("\n");
        nSuccess++;
    }

    /* Prepare report. */
    progressBar.setValue(nMails);
    progressText.setText(tr("Sending Report..."));
    qApp->processEvents();

    QString res = tr("Number of mails: ") + QString::number(nMails) + tr("\n\n") +
                  tr("Mails OK: ") + QString::number(nSuccess) + tr("\n\n") +
                  tr("Mails Failed: ") + QString::number(nFailed) + tr("\n") + failed + tr("\n");

    allTexts.prepend(tr("Beste ") + m_senderName->text() + tr(",\n\n") +
                     tr("Hierbij het rapport van ") + subject + tr("\n\n") +
                     res + tr("\nDe volgende berichten zijn gegenereerd:\n"));
    allTexts.append(tr("\n============================== END ==============================\n"));

    /* Message and content. */
    MimeMessage report;
    MimeText text;

    /* Set sender and receiver. */
    report.setSender(&sender);
    report.addRecipient(&sender);

    /* Add cc's */
    foreach(QString cc, report_cc_addresses){
        report.addCc(new EmailAddress(cc));
    }

    /* Add subject */
    report.setSubject(tr("Report: ") + subject);

    /* Add contents. */
    text.setText(allTexts);
    report.addPart(&text);

    /* Add attachments. */
    foreach(MimeAttachment *att, attachments){
        report.addPart(att);
    }

    if(!sendMail(&report)){
        QMessageBox::warning(this, tr("Error:"), tr("Sending report failed!"));
    }

    progressText.setText(res);
    qApp->processEvents();

    /* Cleanup attachments. */
    foreach(QFile *f, attachmentFiles){
        delete f;
    }
    foreach(MimeAttachment *att, attachments){
        delete att;
    }

    /* Give information. */
    QMessageBox::information(this, tr("Info:"), res);

    progress.hide();

}

/* Wrapper to send an email. */
bool MainWindow::sendMail(MimeMessage *m){

    bool ret = false;

    /* For debugging. */
    if(DO_NOT_SEND_EMAILS){
        this->thread()->sleep(1);
        return false;
    }

    if(m_SMTPConnection == NULL){
        return false;
    }

    try {
        ret = m_SMTPConnection->sendMail(*m);
    }
    catch (...){
        ret = false;
    }

    return ret;

}

/*
 * [9] Dialogs.
 */

/* Create and show about dialog. */
void MainWindow::about(){

    QMessageBox::about(this, tr("About ") + tr(APPLICATION_NAME),
                       tr("<b>A ") + tr(APPLICATION_NAME) + tr("</b>"
                          "<br />"
                          "<br />"
                          "<b>Author:</b><br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") + tr(APPLICATION_AUTHOR) + tr(" (") + tr(APPLICATION_YEAR) + tr(")<br>"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"mailto:") + tr(APPLICATION_AUTHOR_EMAIL) + tr("\">") + tr(APPLICATION_AUTHOR_EMAIL) + tr("</a><br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"http://") + tr(APPLICATION_AUTHOR_URL) + tr(">\">") + tr(APPLICATION_AUTHOR_URL) + tr("</a><br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") + tr(APPLICATION_COMPANY) + tr("<br/>"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Bug reports or issues: <a href=\"mailto:") + tr(APPLICATION_AUTHOR_EMAIL) + tr("\">") + tr(APPLICATION_AUTHOR_EMAIL) + tr("</a><br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;or <a href=\"") + tr(APPLICATION_URL) + tr("\">") + tr(APPLICATION_URL) + tr("</a><br />"
                          "<br />"
                          "<br />"
                          "<b>Credits:</b><br />"
                          "&nbsp;&nbsp; This project makes use of:<br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <a href=\"http://qtxlsx.debao.me/\">Qt Xlsx</a> by <a href=\"https://github.com/dbzhang800\">dbzhang800</a>.<br />"
                          "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <a href=\"https://github.com/bluetiger9/SmtpClient-for-Qt\">SMTP Client for Qt (C++) - Version 1.1</a> by <a href=\"https://github.com/bluetiger9/\">bluetiger9</a>.<br />"
                          "<br />"
                          "<br />"
                          "<b>Note:</b><br />"
                          "This project is in Alpha stage and will most probably contain several bugs and/or stability issues...<br />"
                          "<br /><br /><b>") + tr(APPLICATION_NAME) + tr(" v") + tr(APPLICATION_VERSION) + tr("</b>")
#if defined(__DATE__) && defined(__TIME__)
                          + tr(" (Build: ") + tr(__DATE__) + tr(" ") + tr(__TIME__) + tr(")")
#endif

                        );
}
