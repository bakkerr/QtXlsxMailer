#include "mainwindow.h"

#include <QApplication>

#include <QWidget>
#include <QTableView>

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QGridLayout>

#include <QDialog>
#include <QFileDialog>
#include <QInputDialog>
#include <QMessageBox>

#include <QToolButton>

#include <QRegExp>
#include <QStringRef>

#include <QtXlsx>

#include "xlsxsheetmodel.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent)
{
    this->setWindowTitle(tr("Qt XLSX Email Generator"));

    m_SMTPConnection = NULL;

    setDockNestingEnabled(true);
    setDockOptions(QMainWindow::AnimatedDocks);

    createGeneralOptionsWidget();
    this->addDockWidget(Qt::TopDockWidgetArea, m_generalOptionsDW);

    createXlsxViewerWidget();
    this->addDockWidget(Qt::BottomDockWidgetArea, m_xlsxViewerDW);

    createEditorWidget();
    this->addDockWidget(Qt::LeftDockWidgetArea, m_editorDW);

    createPreviewWidget();
    this->addDockWidget(Qt::RightDockWidgetArea, m_previewDW);

    connect(qApp, SIGNAL(lastWindowClosed()), this, SLOT(saveSettings()));
    loadSettings();

    /* Set default values. */
    updateSheet();

}

/*
 * Destructor.
 *
 * TODO: clean-up properly.
 */
MainWindow::~MainWindow(){

}

void MainWindow::createRowSelectWidget(){

    m_rowSelectWidget = new QFrame(this);
    m_rowSelectWidget->setMinimumWidth(150);
    m_rowSelectWidget->setMaximumWidth(150);

    QVBoxLayout *rowSelectLayout = new QVBoxLayout(m_rowSelectWidget);

    m_firstRowSelect = new QComboBox(m_rowSelectWidget);
    m_firstRowSelect->setToolTip(tr("Select the row where the first email\nshould be generated from."));
    m_lastRowSelect = new QComboBox(m_rowSelectWidget);
    m_lastRowSelect->setToolTip(tr("Select the row where the last email\nshould be generated from."));
    connect(m_firstRowSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));
    connect(m_lastRowSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    m_emailColumnSelect = new QComboBox(m_rowSelectWidget);
    m_emailColumnSelect->setToolTip(tr("Select the column for the\nemail address to use."));
    connect(m_emailColumnSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateText()));
    m_emailAppendText = new QLineEdit(tr("@hr.nl"), m_rowSelectWidget);
    m_emailAppendText->setToolTip(tr("Add a value that should be appended to\n"
                                     "the column where the email address is in.\n"
                                     "If this column already contains a complete\n"
                                     "email address, this field should be empty."));
    connect(m_emailAppendText, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    QPushButton *sendMailsButton = new QPushButton(tr("Send mails"), m_rowSelectWidget);
    connect(sendMailsButton, SIGNAL(clicked()), this, SLOT(sendMails()));

    rowSelectLayout->addWidget(new QLabel(tr("First mail [row]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_firstRowSelect);
    rowSelectLayout->addWidget(new QLabel(tr("Last mail [row]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_lastRowSelect);
    rowSelectLayout->addSpacing(5);
    rowSelectLayout->addWidget(new QLabel(tr("Use the email\naddress in [col]:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_emailColumnSelect);
    rowSelectLayout->addWidget(new QLabel(tr("and append:"), m_rowSelectWidget));
    rowSelectLayout->addWidget(m_emailAppendText);
    rowSelectLayout->addSpacerItem(new QSpacerItem(5, 5, QSizePolicy::Expanding, QSizePolicy::Expanding));
    rowSelectLayout->addWidget(sendMailsButton);

    m_rowSelectWidget->setLayout(rowSelectLayout);
}

void MainWindow::createGeneralOptionsWidget(){

    m_generalOptionsDW = new QDockWidget(tr("General parameters"), this);

    QFrame *generalOptionsWidget = new QFrame(m_generalOptionsDW);
    generalOptionsWidget->setMinimumHeight(130);
    generalOptionsWidget->setMaximumHeight(130);

    QGridLayout *generalOptionsLayout = new QGridLayout(generalOptionsWidget);

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
                              "Multiple addresses may be added seperated by a ';' and\n"
                              "no spaces.\n\n"
                              "Example \"collegue@hr.nl;other@extern.com\""));
    m_emailBcc->setValidator(new QRegExpValidator(QRegExp("(([A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63})[;])*", Qt::CaseInsensitive), m_emailBcc));
    connect(m_emailBcc, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    /* Clear settings button. */
    QPushButton *clearSettingsButton = new QPushButton(tr("Clear Settings"), generalOptionsWidget);
    clearSettingsButton->setToolTip(tr("Reset all saved and edited settings.\nThe application will be as new."));
    connect(clearSettingsButton, SIGNAL(clicked()), this, SLOT(deleteSettings()));

    /* Create the SMTP settings widget. */
    createSMTPWidget();

    /* Add all fields to the layout. */
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Name:"), generalOptionsWidget), 0, 0);
    generalOptionsLayout->addWidget(m_senderName, 0, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Email:"), generalOptionsWidget), 1, 0);
    generalOptionsLayout->addWidget(m_senderEmail, 1, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Add Bcc:"), generalOptionsWidget), 2, 0);
    generalOptionsLayout->addWidget(m_emailBcc, 2, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Email Subject:"), generalOptionsWidget), 0, 2);
    generalOptionsLayout->addWidget(m_emailSubject, 0, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Course Code:"), generalOptionsWidget), 1, 2);
    generalOptionsLayout->addWidget(m_courseCode, 1, 3);
    generalOptionsLayout->addWidget(clearSettingsButton, 2, 3);
    generalOptionsLayout->addWidget(m_SMTPWidget, 0, 4, 3, 1);

    /* Set layout to mainwidget. */
    generalOptionsWidget->setLayout(generalOptionsLayout);

    /* Set mainwidget on DockWidget. */
    m_generalOptionsDW->setWidget(generalOptionsWidget);

}

/* Creates the SMTP settings widget */
void MainWindow::createSMTPWidget(){

    m_SMTPWidget = new QWidget(m_generalOptionsDW);

    QGridLayout *smtpSettingsLayout = new QGridLayout(m_SMTPWidget);
    smtpSettingsLayout->setContentsMargins(0,0,0,0);

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
    m_SMTPWidgetToggleButton->setCheckable(true);
    m_SMTPWidgetToggleButton->setChecked(true);
    m_SMTPWidgetToggleButton->setChecked(false);

    smtpSettingsLayout->addWidget(new QLabel(tr("SMTP server:"), m_SMTPWidget), 0, 1);
    smtpSettingsLayout->addWidget(m_SMTPserver, 0, 2);
    smtpSettingsLayout->addWidget(new QLabel(tr("SMTP port:"), m_SMTPWidget), 1, 1);
    smtpSettingsLayout->addWidget(m_SMTPport, 1, 2);
    smtpSettingsLayout->addWidget(m_SMTPtype, 2, 1);
    smtpSettingsLayout->addWidget(SMTPConnectButton, 2, 2);
    smtpSettingsLayout->addWidget(m_SMTPWidgetToggleButton, 0, 0, 3, 1);

    m_SMTPWidget->setLayout(smtpSettingsLayout);

}

void MainWindow::toggleSMTPWidget(bool s){

    if(!s){
        /* Hide */
        m_toggleSMTPAnimation->setStartValue(400);
        m_toggleSMTPAnimation->setEndValue(20);
        m_SMTPWidgetToggleButton->setText(tr("<"));
        m_SMTPWidgetToggleButton->setToolTip(tr("Show SMTP options"));
    }
    else{
        /* Show. */
        m_toggleSMTPAnimation->setStartValue(20);
        m_toggleSMTPAnimation->setEndValue(400);
        m_SMTPWidgetToggleButton->setText(tr(">"));
        m_SMTPWidgetToggleButton->setToolTip(tr("Hide SMTP options"));
    }

    m_toggleSMTPAnimation->start();

}

void MainWindow::createEditorWidget(){

    /* Create Dockwidget for the editor. */
    m_editorDW = new QDockWidget(tr("Edit:"), this);
    m_editorDW->setMinimumHeight(350);

    /* Create main widget. */
    QFrame *editorWidget = new QFrame(m_editorDW);
    editorWidget->setFrameShape(QFrame::StyledPanel);

    /* Create Layout. */
    QHBoxLayout *editorWidgetLayout = new QHBoxLayout(editorWidget);
    editorWidgetLayout->setContentsMargins(2,2,2,2);

    /* Create tabwidget to contain editors. */
    m_textTab = new QTabWidget(editorWidget);
    m_textTab->setTabPosition(QTabWidget::South);
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
    editorWidgetLayout->addWidget(m_generateWidget);

    /* Set layout to main widget. */
    editorWidget->setLayout(editorWidgetLayout);

    /* Set main widget to dockwidget. */
    m_editorDW->setWidget(editorWidget);

}

void MainWindow::createGenerateWidget(){
    /* Create a frame for this widget. */
    m_generateWidget = new QFrame(m_editorDW);
    m_generateWidget->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::MinimumExpanding);

    /* Create Layout. */
    QGridLayout *generateWidgetLayout = new QGridLayout(m_generateWidget);
    generateWidgetLayout->setContentsMargins(0,0,0,0);

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
    replaceButton->setToolTip(tr("Generate text and overwrite currently selected tab."));
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
    m_generateWidgetToggleButton->setChecked(true);

    /* Add labels to layout */
    generateWidgetLayout->addWidget(new QLabel(tr("Generate email text:"), m_generateWidget), 0, 1, 1, 2);
    generateWidgetLayout->addWidget(new QLabel(tr("Name [col]:"), m_generateWidget), 1, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("Final grade [col]:"), m_generateWidget), 2, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("Include [col]:"), m_generateWidget), 3, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("up to (incl) [col]:"), m_generateWidget), 4, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("Grade title [row]:"), m_generateWidget), 5, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("Max. points [row]:"), m_generateWidget), 6, 1);
    generateWidgetLayout->addWidget(new QLabel(tr("Generate text in:"), m_generateWidget), 8, 1);

    /* Add widgets to layout */
    generateWidgetLayout->addWidget(m_nameColSelect, 1, 2);
    generateWidgetLayout->addWidget(m_finalGradeColSelect, 2, 2);
    generateWidgetLayout->addWidget(m_startColSelect, 3, 2);
    generateWidgetLayout->addWidget(m_stopColSelect, 4, 2);
    generateWidgetLayout->addWidget(m_headerRowSelect, 5, 2);
    generateWidgetLayout->addWidget(m_maxRowSelect, 6, 2);
    generateWidgetLayout->addWidget(newButton, 9, 1);
    generateWidgetLayout->addWidget(replaceButton, 9, 2);

    /* Add the show/hide button. */
    generateWidgetLayout->addWidget(m_generateWidgetToggleButton, 0, 0, 10, 1);

    /* Set layout to main widget. */
    m_generateWidget->setLayout(generateWidgetLayout);
}

/* Show/hide the generateWidget. */
void MainWindow::toggleGenerateWidget(bool s){

    if(!s){
        /* Hide. */
        m_toggleGenerateAnimation->setStartValue(400);
        m_toggleGenerateAnimation->setEndValue(20);
        m_generateWidgetToggleButton->setText(tr("<"));
        m_generateWidgetToggleButton->setToolTip(tr("Show options"));
    }
    else{
        /* Show. */
        m_toggleGenerateAnimation->setStartValue(20);
        m_toggleGenerateAnimation->setEndValue(400);
        m_generateWidgetToggleButton->setText(tr(">"));
        m_generateWidgetToggleButton->setToolTip(tr("Hide options"));
    }

    /* Start animation. */
    m_toggleGenerateAnimation->start();

}

/* Create the selection and preview widget. */
void MainWindow::createPreviewWidget(){

    /* Create the dockwidget. */
    m_previewDW = new QDockWidget(tr("Selection and preview:"), this);

    /* Create a (main)frame for this widget. */
    QFrame *previewWidget = new QFrame(this);
    previewWidget->setFrameShape(QFrame::StyledPanel);
    previewWidget->setMinimumHeight(350);

    /* Create main layout. */
    QHBoxLayout *previewWidgetLayout = new QHBoxLayout();
    previewWidgetLayout->setAlignment(Qt::AlignTop);
    previewWidgetLayout->setContentsMargins(2,2,2,2);

    /* Create layouts. */
    QVBoxLayout *previewBoxLayout = new QVBoxLayout();
    QHBoxLayout *previewSelectionLayout = new QHBoxLayout();

    /* Create 7-segment display for the number of mails. */
    m_nMailsDisplay = new QLCDNumber(4, this);
    m_nMailsDisplay->setFrameStyle(QFrame::NoFrame);
    m_nMailsDisplay->setSegmentStyle(QLCDNumber::Filled);
    m_nMailsDisplay->setPalette(QPalette(Qt::red));
    m_nMailsDisplay->setToolTip(tr("This is the number of emails\nthat this program will send."));

    /* Selection for the row (email) to preview. */
    m_previewSelect = new QComboBox(this);
    m_previewSelect->setToolTip(tr("Select the row (email) you want to preview."));
    connect(m_previewSelect, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    /* The preview tool itself is a read-only textbox. */
    m_previewText = new QTextEdit(this);
    m_previewText->setToolTip(tr("This is how the email looks as it will be send.\n"
                          "You can select an other email in the selection box above."));
    m_previewText->setReadOnly(true);

    /*
     * Create the widget to select the rows to use and the columns for
     * the email address. Optionally you can add a value to the email address.
     */
    createRowSelectWidget();

    /* Set it all in the layouts. */
    previewSelectionLayout->addWidget(new QLabel(tr("Preview:"), this));
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

/* Create the preview tool for xlsx sheets. */
void MainWindow::createXlsxViewerWidget(){

    /* Create the dockwidget that uses all of the available (remaining) space. */
    m_xlsxViewerDW = new QDockWidget(tr("XLSX Viewer:"), this);
    m_xlsxViewerDW->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::MinimumExpanding);

    /* Create a frame for this widget that uses all of the available (remaining) space. */
    QFrame *xlsxWidget = new QFrame(this);
    xlsxWidget->setMinimumHeight(150);
    xlsxWidget->setMinimumWidth(400);
    xlsxWidget->setSizePolicy(QSizePolicy::MinimumExpanding, QSizePolicy::MinimumExpanding);

    QVBoxLayout *l = new QVBoxLayout();
    QHBoxLayout *l1 = new QHBoxLayout();

    m_loadXlsxFileButton = new QPushButton(tr("Load xlsx file"), this);
    connect(m_loadXlsxFileButton, SIGNAL(clicked()), this, SLOT(loadSheet()));

    /* Add the tabwidget where the tabs from the xlsx file can be loaded. */
    m_xlsxTab = new QTabWidget(this);
    m_xlsxTab->setTabPosition(QTabWidget::South);
    m_xlsxTab->setTabsClosable(true);
    connect(m_xlsxTab, SIGNAL(tabCloseRequested(int)), this, SLOT(closeTab(int)));
    connect(m_xlsxTab, SIGNAL(currentChanged(int)), this, SLOT(updateSheet()));

    /* Set layout. */
    l1->addStretch(1);
    l1->addWidget(m_loadXlsxFileButton);
    l->addLayout(l1);
    l->addWidget(m_xlsxTab);

    /* Set layout to main widget. */
    xlsxWidget->setLayout(l);

    /* Set main widget to dockwidget. */
    m_xlsxViewerDW->setWidget(xlsxWidget);

}

/* Todo: Place in toolbar */
QWidget* MainWindow::createInfoBar(){

    QWidget *infoWidget = new QWidget(this);

    QHBoxLayout *l = new QHBoxLayout();

    m_extraText = new QLabel(this);
    m_extraText->setText(tr(APPLICATION_AUTHOR) + tr(" - ")+ tr(APPLICATION_COMPANY) + tr(" - ") + tr(APPLICATION_YEAR));

    QPushButton *help = new QPushButton(tr("Help"), this);
    help->setEnabled(false);

    QPushButton *about = new QPushButton(tr("About"), this);
    connect(about, SIGNAL(clicked()), this, SLOT(about()));

    QPushButton *qtInfo = new QPushButton(tr("Qt"), this);
    connect(qtInfo, SIGNAL(clicked()), qApp, SLOT(aboutQt()));

    l->addWidget(m_extraText);
    l->addWidget(help);
    l->addWidget(about);
    l->addWidget(qtInfo);

    infoWidget->setLayout(l);

    return infoWidget;
}

/* Load an xlsx sheet */
void MainWindow::loadSheet(){

    /* Ask for the file. */
    QString filePath = QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), "*.xlsx");
    if(filePath.isEmpty()){
        return;
    }

    /* Open the document. */
    QXlsx::Document *xlsx = new QXlsx::Document(filePath);

    /* Add add sheets (tabs) to the viewer. */
    foreach(QString sheetName, xlsx->sheetNames()){

        /* Load sheet */
        QXlsx::Worksheet *sheet = dynamic_cast<QXlsx::Worksheet *>(xlsx->sheet(sheetName));
        if(sheet){

            /* Create a tableview for this sheet. */
            QTableView *view = new QTableView(m_xlsxTab);

            /* Set to read-only. */
            view->setEditTriggers(QAbstractItemView::NoEditTriggers);
            view->setModel(new QXlsx::SheetModel(sheet, view));

            /* Handle merged cells. */
            foreach (QXlsx::CellRange range, sheet->mergedCells()){
                view->setSpan(range.firstRow()-1, range.firstColumn()-1, range.rowCount(), range.columnCount());
            }

            /* Add sheet as a tab to viewer. */
            m_xlsxTab->addTab(view, sheetName);
        }
    }
}

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

/* Hacky thing to avoid circular updates when loading new values into the selection boxes. */
void MainWindow::blockRowSignals(bool b){
    m_firstRowSelect->blockSignals(b);
    m_lastRowSelect->blockSignals(b);
    m_previewSelect->blockSignals(b);
}

/* Slot called when selecting an onther sheet. */
void MainWindow::updateSheet(){

    /* Get a pointer to the data. */
    QTableView *d = (QTableView*)m_xlsxTab->currentWidget();

    /* Extract columns and rows. */
    QStringList columnNames;
    QStringList rowNames;
    if(d != NULL){
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

    QComboBox *cols[] = {m_nameColSelect, m_finalGradeColSelect, m_startColSelect, m_stopColSelect};
    for(int i = 0; i < 4; i++){
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

    if(d != NULL){
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
        if(i >= start && i <= stop){
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
        m_nMailsDisplay->display(m_lastRowSelect->currentText().toInt() - m_firstRowSelect->currentText().toInt() + 1);
    }
    else{
        m_nMailsDisplay->display(0);
    }

    /* Update preview text. */
    updateText();
}

/* Update and parse values. */
void MainWindow::updateText(){

    /* Check option fields and make red when useless. */
    !isValidEmail(m_senderEmail->text()) ? m_senderEmail->setStyleSheet(tr("background-color: #FF9999;")) :
                                           m_senderEmail->setStyleSheet(tr(""));

    m_emailSubject->text().length() < 3  ? m_emailSubject->setStyleSheet(tr("background-color: #FF9999;")) :
                                           m_emailSubject->setStyleSheet(tr(""));

    m_courseCode->text().length() < 3    ? m_courseCode->setStyleSheet(tr("background-color: #FF9999;")) :
                                           m_courseCode->setStyleSheet(tr(""));

    m_xlsxTab->count() < 1               ? m_loadXlsxFileButton->setStyleSheet(tr("background-color: #FF9999;")) :
                                           m_loadXlsxFileButton->setStyleSheet(tr(""));

    QStringList bcc_addresses = m_emailBcc->text().split(";");
    foreach(QString bcc, bcc_addresses){
        !bcc.isEmpty() && !isValidEmail(bcc) ? m_emailBcc->setStyleSheet(tr("background-color: #FF9999;")) :
                                               m_emailBcc->setStyleSheet(tr(""));
    }

    /* Generate preview text. */
    QString res;

    /* Get the mail number. */
    int offset = m_previewSelect->currentText().toInt();

    /* Extract default fields. */
    res += tr("From: ") + m_senderName->text() + tr(" <") + m_senderEmail->text() +  tr(">\n");
    res += tr("To: <") + getData(m_emailColumnSelect->currentText(), offset) + m_emailAppendText->text() + tr(">\n");
    foreach(QString bcc, bcc_addresses){
        if(!bcc.isEmpty()){
            res += tr("Bcc: <") + bcc + tr(">\n");
        }
    }
    res += tr("Subject: [") + m_courseCode->text() + tr("] ") + m_emailSubject->text() + tr("\n");
    res += tr("\n\n");

    /* Get the main contents. */
    res += getMailText(offset);

    /* set text to preview. */
    m_previewText->setText(res);
}

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

/* The main thing.. Sending emails. */
void MainWindow::sendMails(){

    /* Read start/stop rows and calculate number of mails. */
    int start = m_firstRowSelect->currentText().toInt();
    int stop = m_lastRowSelect->currentText().toInt();
    int nMails = stop - start + 1;

    /* Check sender. */
    QString fromEmail = m_senderEmail->text();
    if(!isValidEmail(fromEmail)){
        QMessageBox::warning(this, tr("Error:"), tr("Sender email address is invalid!"));
        m_senderEmail->setFocus();
        return;
    }

    /* Use name or email address as sender. */
    QString fromName = m_senderName->text();
    if(fromName.isEmpty()){
        fromName = fromEmail;
    }

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

    /* Check is there are any mails to send. */
    if(m_nMailsDisplay->value() == 0){
        QMessageBox::warning(this, tr("Error:"), tr("The number of mails is 0!"));
        return;
    }

    /* Check mail addresses and contents. */
    for(int i = start; i <= stop; i++){
        QString recv_mail = getData(m_emailColumnSelect->currentText(), i) + m_emailAppendText->text();

        /* Recipient address OK? */
        if(!isValidEmail(recv_mail)){
            QMessageBox::warning(this, tr("Error:"), tr("The email address ") + recv_mail + tr(" on line ") + QString::number(i) + tr(" is invalid!"));
            m_emailColumnSelect->setFocus();
            return;
        }

        /* Mailtext OK? */
        if(getMailText(i).contains("[INV_REF!]")){
            QMessageBox::warning(this, tr("Error:"), tr("There are invalid references in the mailtext of email ") + QString::number(i) + tr("!"));
            return;
        }
    }

    /* Get bcc addresses. */
    QStringList bcc_addresses = m_emailBcc->text().split(";");

    /* Check if they are OK. */
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

    /* Do we already have a connection? If not, connect. */
    if(m_SMTPConnection == NULL){
        SMTPconnect();
        if(m_SMTPConnection == NULL){
            return;
        }
    }

    /* Sure? */
    if(QMessageBox::question(this, tr("Send Emails now?"),
                                   tr("Are you sure you want to send ") +
                                   QString::number(m_nMailsDisplay->value()) +
                                   tr(" emails now ?")
                             ) != QMessageBox::Yes){
        return;
    }

    QString success;
    int nSuccess = 0;
    QString failed;
    int nFailed = 0;

    /* Sender is the same for each mail. */
    EmailAddress sender(fromEmail, fromName);

    /* Generate and send messages. */
    for(int i = start; i <= stop; i++){

        /* Message and content. */
        MimeMessage message;
        MimeText text;

        /* Receiver email address. */
        QString recv_mail = getData(m_emailColumnSelect->currentText(), i) + m_emailAppendText->text();
        EmailAddress receiver(recv_mail);

        /* Set sender and receiver. */
        message.setSender(&sender);
        message.addRecipient(&receiver);

        /* Add bcc's */
        foreach(QString bcc, bcc_addresses){
            message.addBcc(new EmailAddress(bcc));
        }

        /* Add subject */
        message.setSubject(subject);

        /* Add contents. */
        text.setText(getMailText(i));

        /* Add content to message. */
        message.addPart(&text);

        /* Try to send the mail. If failed, keep track of this. */
        try {
            if(m_SMTPConnection == NULL || !m_SMTPConnection->sendMail(message)){
                failed += tr("  ") + recv_mail + tr("\n");
                nFailed++;
                continue;
            }

            success += tr("  ") + recv_mail + tr("\n");
            nSuccess++;
        }
        catch (...){
            failed += tr("  ") + recv_mail + tr("\n");
            nFailed++;
        }

    }

    /* Give information. */
    QMessageBox::information(this, tr("Info:"),
                                   tr("Tried to send ") + QString::number(nMails) + tr(" mails.\n\n") +
                                   tr("Mails OK: ") + QString::number(nSuccess) + tr("\n\n") +
                                   tr("Mails Failed:") + QString::number(nFailed) + tr("\n") + failed + tr("\n")
                             );
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
    if(row < 0 || col < 0 || d == NULL || d->model() == NULL || col > d->model()->columnCount() || row > d->model()->rowCount()){
        return QString("[INV_REF!]");
    }

    /* Return the value at row,col */
    return(d->model()->data(d->model()->index(row-1, col-1)).toString());
}

/* Add new composer tab and make it active. */
void MainWindow::addNewTextTab(){

    QTextEdit *newText = new QTextEdit(this);
    connect(newText, SIGNAL(textChanged()), this, SLOT(updateText()));

    int num = m_textTab->count() - 1;
    int newNum = 1;
    for(int i = num; i >= 0; i--){
        QTextEdit *te;
        if((te = qobject_cast<QTextEdit*>(m_textTab->widget(i)))){
            QRegExp re("[T](\\d+)");
            QString name = m_textTab->tabText(i);
            qDebug() << name << "..." << endl;
            if(re.exactMatch(name)){
                int oldNum = re.cap(1).toInt();
                qDebug() << "match " << name << re.cap(0) << re.cap(1) << oldNum << endl;
                if(oldNum >= newNum){
                    newNum = oldNum + 1;
                }
            }
        }
    }

    m_textTab->addTab(newText, tr("T") + QString::number(newNum));
    m_textTab->setCurrentWidget(newText);

}

/* Rename a tab in the editor. */
void MainWindow::renameTab(int index){

    /* Get the calling tabWidget. */
    QTabWidget *tw = (QTabWidget*)this->sender();

    /* Ask for new name. */
    bool ok;
    QString oldname = tw->tabText(index);
    QString name = QInputDialog::getText(this, tr("Set new name for tab:"), tr("New name:"), QLineEdit::Normal, oldname, &ok);

    /* If ok, then rename. */
    if(ok){
        tw->setTabText(index, name);
    }
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

        /* Include maxpoints/default? */
        if(maxpoints){
            txt += tr("/#") + m_finalGradeColSelect->currentText() + m_maxRowSelect->currentText() + tr("#");
        }

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

    /* Create new tab? */
    if(newTab){
        addNewTextTab();
    }

    /* Set text. */
    ((QTextEdit*)m_textTab->currentWidget())->setText(txt);

    /* We're done! Hide the generate widget. */
    m_generateWidgetToggleButton->setChecked(false);

}

/* Validate email address. */
bool MainWindow::isValidEmail(QString address){
    return QRegExp("[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z0-9-]{2,63}", Qt::CaseInsensitive).exactMatch(address);
}

void MainWindow::saveSettings(){

    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR) + tr(APPLICATION_VERSION), this);

    s->setValue(tr("senderName"), m_senderName->text());
    s->setValue(tr("senderEmail"), m_senderEmail->text());
    s->setValue(tr("emailBcc"), m_emailBcc->text());
    s->setValue(tr("emailSubject"), m_emailSubject->text());
    s->setValue(tr("emailAppend"), m_emailAppendText->text());
    s->setValue(tr("courseCode"), m_courseCode->text());
    s->setValue(tr("SMTPserver"), m_SMTPserver->text());
    s->setValue(tr("SMTPport"), m_SMTPport->text());
    s->setValue(tr("SMTPtype"), m_SMTPtype->currentText());

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

    s->sync();

}

void MainWindow::loadSettings(){

    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR) + tr(APPLICATION_VERSION), this);

    m_senderName->setText(s->value(tr("senderName"), tr("")).toString());
    m_senderEmail->setText(s->value(tr("senderEmail"), tr("")).toString());
    m_emailBcc->setText(s->value(tr("emailBcc"), tr("")).toString());
    m_emailSubject->setText(s->value(tr("emailSubject"), tr("")).toString());
    m_emailAppendText->setText(s->value(tr("emailAppend"), tr("@hr.nl")).toString());
    m_courseCode->setText(s->value(tr("courseCode"), tr("")).toString());
    m_SMTPserver->setText(s->value(tr("SMTPserver"), tr("smtp.hr.nl")).toString());
    m_SMTPport->setText(s->value(tr("SMTPport"), tr("465")).toString());
    m_SMTPtype->setCurrentText(s->value(tr("SMTPtype"), tr("SSL")).toString());

    int num = s->beginReadArray(tr("mailTexts"));
    for(int i = 0; i < num; i++){
        s->setArrayIndex(i);
        QString name = s->value(tr("name"), tr("unknown")).toString();
        addNewTextTab();
        qDebug() << i << name << m_textTab->currentIndex() << endl;
        m_textTab->setTabText(m_textTab->currentIndex(), name);
        ((QTextEdit*)m_textTab->currentWidget())->setText(s->value(tr("text"), tr("")).toString());
    }
    if(num == 0){
        addNewTextTab();
    }
    s->endArray();



}

void MainWindow::deleteSettings(){

    QSettings *s = new QSettings(tr(APPLICATION_COMPANY_ABBR), tr(APPLICATION_NAME_ABBR) + tr(APPLICATION_VERSION), this);

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
