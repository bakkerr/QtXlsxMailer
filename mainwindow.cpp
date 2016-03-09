#include "mainwindow.h"

#include <QApplication>

#include <QWidget>
#include <QTableView>
#include <QPushButton>

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QGridLayout>

#include <QDialog>
#include <QFileDialog>
#include <QInputDialog>
#include <QMessageBox>

#include <QRegExp>
#include <QStringRef>

#include <QtXlsx>

#include "xlsxsheetmodel.h"


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent)
{
    this->setWindowTitle(tr("Qt XLSX Email Generator"));

    m_smtp = NULL;

    QWidget *mainWidget = new QWidget(this);
    mainWidget->setMinimumSize(650, 700);

    QVBoxLayout *l1 = new QVBoxLayout();

    l1->addWidget(createGeneralOptionsWidget());
    l1->addWidget(createMiddleWidget());
    l1->addWidget(createXlsxViewer());
    l1->addWidget(createInfoBar());

    mainWidget->setLayout(l1);

    this->setCentralWidget(mainWidget);

    updateText();

}

MainWindow::~MainWindow()
{

}


QWidget* MainWindow::createRowSelectWidget(){

    QFrame *rowSelectWidget = new QFrame(this);
    rowSelectWidget->setFrameShape(QFrame::Box);
    rowSelectWidget->setMinimumWidth(150);
    rowSelectWidget->setMaximumWidth(150);
    rowSelectWidget->setMinimumHeight(350);

    QVBoxLayout *rowSelectLayout = new QVBoxLayout();

    QHBoxLayout *numLayout = new QHBoxLayout();
    m_nMails = new QLCDNumber(4, this);
    m_nMails->setFrameStyle(QFrame::NoFrame);
    m_nMails->setSegmentStyle(QLCDNumber::Filled);
    m_nMails->setPalette(QPalette(Qt::red));
    numLayout->addWidget(new QLabel(tr("Mails:"), this));
    numLayout->addWidget(m_nMails);

    m_start = new QComboBox(this);
    m_stop = new QComboBox(this);
    connect(m_start, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));
    connect(m_stop, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    m_emailColumn = new QComboBox(this);
    connect(m_emailColumn, SIGNAL(currentTextChanged(QString)), this, SLOT(updateText()));
    m_emailAppend = new QLineEdit(tr("@hr.nl"), this);
    connect(m_emailAppend, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    QPushButton *sendMailsButton = new QPushButton(tr("Send mails"), this);
    connect(sendMailsButton, SIGNAL(clicked()), this, SLOT(sendMails()));

    rowSelectLayout->addLayout(numLayout);
    rowSelectLayout->addSpacing(20);
    rowSelectLayout->addWidget(new QLabel(tr("Start from row:"), this));
    rowSelectLayout->addWidget(m_start);
    rowSelectLayout->addWidget(new QLabel(tr("to row (included):"), this));
    rowSelectLayout->addWidget(m_stop);
    rowSelectLayout->addSpacing(20);
    rowSelectLayout->addWidget(new QLabel(tr("Use the address\nin Column:"), this));
    rowSelectLayout->addWidget(m_emailColumn);
    rowSelectLayout->addWidget(new QLabel(tr("and append:"), this));
    rowSelectLayout->addWidget(m_emailAppend);
    rowSelectLayout->addSpacerItem(new QSpacerItem(20, 20, QSizePolicy::Expanding, QSizePolicy::Expanding));
    rowSelectLayout->addWidget(sendMailsButton);

    rowSelectWidget->setLayout(rowSelectLayout);

    return rowSelectWidget;
}

QWidget* MainWindow::createGeneralOptionsWidget(){

    QFrame *generalOptionsWidget = new QFrame(this);
    generalOptionsWidget->setFrameShape(QFrame::Box);

    QGridLayout *generalOptionsLayout = new QGridLayout();

    m_senderName = new QLineEdit(tr(""), this);
    connect(m_senderName, SIGNAL(textChanged(QString)), this, SLOT(updateText()));
    m_senderEmail = new QLineEdit(tr(""), this);
    connect(m_senderEmail, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    m_emailSubject = new QLineEdit(tr(""), this);
    connect(m_emailSubject, SIGNAL(textChanged(QString)), this, SLOT(updateText()));
    m_courseCode = new QLineEdit(tr(""), this);
    connect(m_courseCode, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    m_emailBcc = new QLineEdit(tr(""), this);
    connect(m_emailBcc, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    m_SMTPserver = new QLineEdit(tr("smtp.hr.nl"), this);
    m_SMTPport = new QLineEdit(tr("465"), this);
    m_SMTPport->setValidator(new QIntValidator(1, 65535, this));

    m_SMTPtype = new QComboBox(this);
    m_SMTPtype->addItem(tr("SSL"), SmtpClient::SslConnection);
    m_SMTPtype->addItem(tr("TLS"), SmtpClient::TlsConnection);
    m_SMTPtype->addItem(tr("TCP"), SmtpClient::TcpConnection);
    m_SMTPtype->setCurrentText(tr("SSL"));

    QPushButton *SMTPConnectButton = new QPushButton(tr("SMTP Connect"), this);
    connect(SMTPConnectButton, SIGNAL(clicked()), this, SLOT(SMTPconnect()));

    generalOptionsLayout->addWidget(new QLabel(tr("Sender Name:"), this), 0, 0);
    generalOptionsLayout->addWidget(m_senderName, 0, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Email:"), this), 1, 0);
    generalOptionsLayout->addWidget(m_senderEmail, 1, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Add Bcc:"), this), 2, 0);
    generalOptionsLayout->addWidget(m_emailBcc, 2, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Email Subject:"), this), 0, 2);
    generalOptionsLayout->addWidget(m_emailSubject, 0, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("Course Code:"), this), 1, 2);
    generalOptionsLayout->addWidget(m_courseCode, 1, 3);
    generalOptionsLayout->addWidget(new QLabel(tr("SMTP server:"), this), 0, 4);
    generalOptionsLayout->addWidget(m_SMTPserver, 0, 5, 1, 2);
    generalOptionsLayout->addWidget(new QLabel(tr("SMTP port:"), this), 1, 4);
    generalOptionsLayout->addWidget(m_SMTPport, 1, 5, 1, 2);
    generalOptionsLayout->addWidget(m_SMTPtype, 2, 5);
    generalOptionsLayout->addWidget(SMTPConnectButton, 2, 6);

    generalOptionsWidget->setLayout(generalOptionsLayout);

    return generalOptionsWidget;
}

QWidget* MainWindow::createEditorWidget(){

    QFrame *editorWidget = new QFrame(this);
    editorWidget->setFrameShape(QFrame::Box);
    editorWidget->setMinimumHeight(350);

    QVBoxLayout *l1 = new QVBoxLayout();
    QHBoxLayout *l2 = new QHBoxLayout();

    m_edit = new QTextEdit(this);
    m_edit->setText(tr(""));
    connect(m_edit, SIGNAL(textChanged()), this, SLOT(updateText()));

    QPushButton *generateButton = new QPushButton(tr("Generate"), this);
    connect(generateButton, SIGNAL(clicked()), this, SLOT(generateTextDialog()));

    l2->addWidget(new QLabel(tr("Edit:")));
    l2->addWidget(generateButton);

    l1->addLayout(l2);
    l1->addWidget(m_edit);

    editorWidget->setLayout(l1);

    return editorWidget;
}

QWidget* MainWindow::createPreviewWidget(){

    QFrame *previewWidget = new QFrame(this);
    previewWidget->setFrameShape(QFrame::Box);
    previewWidget->setMinimumHeight(350);

    QVBoxLayout *l1 = new QVBoxLayout();
    QHBoxLayout *l2 = new QHBoxLayout();

    m_preview = new QComboBox(this);
    connect(m_preview, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    m_view = new QTextEdit(this);
    m_view->setReadOnly(true);

    l2->addWidget(new QLabel(tr("Preview:")));
    l2->addWidget(m_preview);
    l1->addLayout(l2);
    l1->addWidget(m_view);

    previewWidget->setLayout(l1);

    return previewWidget;
}

QWidget* MainWindow::createMiddleWidget(){

    QWidget *middleWidget = new QWidget();
    middleWidget->setContentsMargins(0,0,0,0);
    middleWidget->setMinimumHeight(350);

    QHBoxLayout *l = new QHBoxLayout();
    l->setAlignment(Qt::AlignTop);
    l->setContentsMargins(0,0,0,0);

    l->addWidget(createEditorWidget());
    l->addWidget(createRowSelectWidget());
    l->addWidget(createPreviewWidget());

    middleWidget->setLayout(l);

    return middleWidget;
}

QWidget* MainWindow::createXlsxViewer(){

    QFrame *xlsxWidget = new QFrame(this);
    xlsxWidget->setFrameShape(QFrame::Box);
    xlsxWidget->setMinimumHeight(150);
    xlsxWidget->setMinimumWidth(400);

    QVBoxLayout *l = new QVBoxLayout();
    QHBoxLayout *l1 = new QHBoxLayout();

    QPushButton *loadXlsxButton = new QPushButton(tr("Load xlsx file"), this);
    connect(loadXlsxButton, SIGNAL(clicked()), this, SLOT(loadSheet()));

    m_tabWidget = new QTabWidget(this);
    m_tabWidget->setTabPosition(QTabWidget::South);
    m_tabWidget->setTabsClosable(true);
    connect(m_tabWidget, SIGNAL(tabCloseRequested(int)), this, SLOT(closeTab(int)));
    connect(m_tabWidget, SIGNAL(currentChanged(int)), this, SLOT(updateSheet()));

    l1->addWidget(new QLabel(tr("XLSX Viewer:")));
    l1->addWidget(loadXlsxButton);
    l->addLayout(l1);
    l->addWidget(m_tabWidget);

    xlsxWidget->setLayout(l);

    return xlsxWidget;
}

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

void MainWindow::loadSheet(){

    QString filePath = QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), "*.xlsx");
    if(filePath.isEmpty()){
        return;
    }

    QXlsx::Document *xlsx = new QXlsx::Document(filePath);
    foreach(QString sheetName, xlsx->sheetNames()){
        QXlsx::Worksheet *sheet = dynamic_cast<QXlsx::Worksheet *>(xlsx->sheet(sheetName));
        if(sheet){
            QTableView *view = new QTableView(m_tabWidget);
            view->setModel(new QXlsx::SheetModel(sheet, view));
            foreach (QXlsx::CellRange range, sheet->mergedCells()){
                view->setSpan(range.firstRow()-1, range.firstColumn()-1, range.rowCount(), range.columnCount());
            }
            m_tabWidget->addTab(view, sheetName);
        }
    }
}

void MainWindow::closeTab(int index){
    m_tabWidget->removeTab(index);
}

void MainWindow::blockRowSignals(bool b){
    m_start->blockSignals(b);
    m_stop->blockSignals(b);
    m_preview->blockSignals(b);
}

void MainWindow::updateSheet(){
    QTableView *d = (QTableView*)m_tabWidget->currentWidget();

    m_emailColumn->clear();

    QStringList headers;
    if(d != NULL){
        for(int i = 0; i < d->model()->columnCount(); i++){
          headers.append(d->model()->headerData(i, Qt::Horizontal).toString());
        }
    }

    m_emailColumn->addItems(headers);

    updateInfo();
}

void MainWindow::updateInfo(){
    blockRowSignals(true);

    QTableView *d = (QTableView*)m_tabWidget->currentWidget();
    int max = 0;
    int start = 1;
    int stop = max;
    int preview = 1;

    if(d != NULL){
        max = d->model()->rowCount();
        stop = max;
    }

    if(!m_start->currentText().isEmpty()){
        start = m_start->currentText().toInt();
    }
    if(!m_stop->currentText().isEmpty()){
        stop = m_stop->currentText().toInt();
    }
    if(!m_preview->currentText().isEmpty()){
        preview = m_preview->currentText().toInt();
    }

    m_start->clear();
    m_stop->clear();
    m_preview->clear();

    for(int i = 1; i <= max; i++){
        m_start->addItem(QString::number(i));
        if(i >= start && i <= stop){
            m_preview->addItem(QString::number(i));
        }
        if(i >= start){
            m_stop->addItem(QString::number(i));
        }
    }

    m_start->setCurrentText(QString::number(start));
    if(stop >= start){
        m_stop->setCurrentText(QString::number(stop));
    }
    else{
        m_stop->setCurrentText(QString::number(max));
    }
    m_preview->setCurrentText(QString::number(preview));

    blockRowSignals(false);

    if(max > 0){
        m_nMails->display(m_stop->currentText().toInt() - m_start->currentText().toInt() + 1);
    }
    else{
        m_nMails->display(0);
    }

    updateText();
}

void MainWindow::updateText(){
    QString res;

    int offset = m_preview->currentText().toInt();

    res += tr("From: ") + m_senderName->text() + tr(" <") + m_senderEmail->text() +  tr(">\n");
    res += tr("To: <") + getData(m_emailColumn->currentText(), offset) + m_emailAppend->text() + tr(">\n");
    res += tr("Bcc: <") + m_emailBcc->text() + tr(">\n");
    res += tr("Subject: [") + m_courseCode->text() + tr("] ") + m_emailSubject->text() + tr("\n");
    res += tr("\n\n");
    res += getMailText(offset);
    m_view->setText(res);
}

void MainWindow::SMTPconnect(){
    bool ok;

    QString smtpServer = m_SMTPserver->text();
    int smtpPort = m_SMTPport->text().toInt();
    SmtpClient::ConnectionType smtpType = static_cast<SmtpClient::ConnectionType>(m_SMTPtype->currentData().toInt());

    try {
        m_smtp = new SmtpClient(smtpServer, smtpPort, smtpType);
    }
    catch (...){

    }

    QString user = QInputDialog::getText(this, tr("Username:"),
                                           tr("SMTP username for ") + smtpServer +
                                           tr(":") + QString::number(smtpPort),
                                           QLineEdit::Normal, m_senderEmail->text(),
                                         &ok);

    if(!ok || user.isEmpty()){
        SMTPdisconnect();
        return;
    }

    m_smtp->setUser(user);

    QString password = QInputDialog::getText(this, tr("Password:"),
                                               tr("SMTP Password for <") + user +
                                               tr(">@") + smtpServer +
                                               tr(":") + QString::number(smtpPort),
                                               QLineEdit::Password, tr(""),
                                             &ok);

    if(!ok || password.isEmpty()){
        SMTPdisconnect();
        return;
    }

    m_smtp->setPassword(password);

    try {
        if(!m_smtp->connectToHost()){
            QMessageBox::warning(this, tr("SMTP Connect"), tr("Could not connect to SMTP server!"));
            SMTPdisconnect();
            return;
        }

        if(!m_smtp->login()){
            QMessageBox::warning(this, tr("SMTP Connect"), tr("SMTP login failed!"));
            SMTPdisconnect();
            return;
        }
    }
    catch (...){

    }

}

void MainWindow::SMTPdisconnect(){
    if(m_smtp != NULL){
    /*    try {
            m_smtp->quit();
        }
        catch (...){

    */
    delete m_smtp;
    }

    m_smtp = NULL;
}

void MainWindow::sendMails(){

    int start = m_start->currentText().toInt();
    int stop = m_stop->currentText().toInt();
    int nMails = stop - start + 1;

    QString fromEmail = m_senderEmail->text();
    if(fromEmail.isEmpty()){
        QMessageBox::warning(this, tr("Error:"), tr("Sender email cannot be empty!"));
        m_senderEmail->setFocus();
        return;
    }

    QString fromName = m_senderName->text();
    if(fromName.isEmpty()){
        fromName = fromEmail;
    }

    QString coursecode = m_courseCode->text();
    if(coursecode.length() < 3){
        QMessageBox::warning(this, tr("Error:"), tr("Course code cannot be less than 3 characters!"));
        m_courseCode->setFocus();
        return;
    }

    QString subject = tr("[") + coursecode + tr("] ") + m_emailSubject->text();
    if(m_emailSubject->text().isEmpty()){
        QMessageBox::warning(this, tr("Error:"), tr("Subject cannot be empty!"));
        m_emailSubject->setFocus();
        return;
    }

    QString bcc = m_emailBcc->text();

    if(m_nMails->value() == 0){
        QMessageBox::warning(this, tr("Error:"), tr("The number of mails is 0!"));
        return;
    }

    for(int i = start; i <= stop; i++){
        if(getMailText(i).contains("[INV_REF!]")){
            QMessageBox::warning(this, tr("Error:"), tr("There are invalid references in the mailtext!"));
            m_edit->setFocus();
            return;
        }
    }

    if(m_smtp == NULL){
        SMTPconnect();
        if(m_smtp == NULL){
            return;
        }
    }

    if(QMessageBox::question(this, tr("Send Emails now?"),
                                   tr("Are you sure you want to send ") +
                                   QString::number(m_nMails->value()) +
                                   tr(" emails now ?")
                             ) != QMessageBox::Yes){
        return;
    }

    EmailAddress sender(fromEmail, fromName);
    EmailAddress bcc_addr(bcc);

    QString success;
    int nSuccess = 0;
    QString failed;
    int nFailed = 0;

    for(int i = start; i <= stop; i++){
        MimeMessage message;
        MimeText text;

        QString recv_mail = getData(m_emailColumn->currentText(), i) + m_emailAppend->text();
        EmailAddress receiver(recv_mail);

        message.setSender(&sender);
        message.addRecipient(&receiver);
        if(!bcc.isEmpty()){
            message.addBcc(&bcc_addr);
        }

        message.setSubject(subject);

        text.setText(getMailText(i));

        message.addPart(&text);

        try {
            if(m_smtp == NULL || !m_smtp->sendMail(message)){
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

    QMessageBox::information(this, tr("Info:"),
                                   tr("Tried to send ") + QString::number(nMails) + tr(" mails.\n\n") +
                                   tr("Mails OK: ") + QString::number(nSuccess) + tr("\n\n") +
                                   tr("Mails Failed:") + nFailed + tr("\n") + failed + tr("\n")
                             );
}

QString MainWindow::getMailText(int offset){
    QString txt = m_edit->toPlainText();
    QRegExp re("#([A-Z,a-z,0-9]*)#");

    QString res;

    QStringList list;
    int lastpos = 0;
    int pos = 0;
    while ((pos = re.indexIn(txt, pos)) != -1) {
        res.append(txt.mid(lastpos, pos-lastpos));
        QString cell = re.cap(1);
        res.append(getData(cell, offset));
        list << cell;
        pos += re.matchedLength();
        lastpos = pos;
    }

    res.append(txt.mid(lastpos, txt.length()-lastpos));

    return res;
}

QString MainWindow::getData(QString cell, int offset)
{
    QRegExp re("([A-Z,a-z]+)([0-9]*)");

    if(re.indexIn(cell, 0) != -1){
        QString col = re.cap(1).toUpper();
        QString row = re.cap(2);

        int r;
        if(row.isEmpty()){
            r = offset;
        }
        else{
            r = row.toInt();
        }

        int c = 0;

        switch(col.length()){
          case 4:
            c = (qPow(26, 3) * (col.data()[0].toLatin1() - 'A' + 1)) +
                (qPow(26, 2) * (col.data()[1].toLatin1() - 'A' + 1)) +
                (qPow(26, 1) * (col.data()[2].toLatin1() - 'A' + 1)) +
                (qPow(26, 0) * (col.data()[3].toLatin1() - 'A' + 1));
            break;
          case 3:
            c = (qPow(26, 2) * (col.data()[0].toLatin1() - 'A' + 1)) +
                (qPow(26, 1) * (col.data()[1].toLatin1() - 'A' + 1)) +
                (qPow(26, 0) * (col.data()[2].toLatin1() - 'A' + 1));
            break;
          case 2:
            c = (qPow(26, 1) * (col.data()[0].toLatin1() - 'A' + 1)) +
                (qPow(26, 0) * (col.data()[1].toLatin1() - 'A' + 1));
            break;
          case 1:
            c = col.data()[0].toLatin1() - 'A' + 1;
            break;
          default:
            c = 0;
        }

        return getData(r, c);
    }

    return QString("[INV_REF!]");
}

QString MainWindow::getData(int row, int column)
{

    QTableView *d = (QTableView*)m_tabWidget->currentWidget();

    if(row < 0 || column < 0 || d == NULL || d->model() == NULL || column > d->model()->columnCount() || row > d->model()->rowCount()){
        return QString("[INV_REF!]");
    }

    return(d->model()->data(d->model()->index(row-1, column-1)).toString());
}

void MainWindow::generateTextDialog(){

    QTableView *d = (QTableView*)m_tabWidget->currentWidget();
    QComboBox nameCol;
    QComboBox finalCol;
    QComboBox startCol;
    QComboBox stopCol;
    QComboBox maxRow;
    QComboBox headerRow;

    QStringList columnNames;
    QStringList rowNames;

    columnNames.append(tr("<Don't include>"));
    rowNames.append(tr("<Don't include>"));

    if(d != NULL){
        for(int i = 0; i < d->model()->columnCount(); i++){
            columnNames.append(d->model()->headerData(i, Qt::Horizontal).toString());
        }
        for(int i = 1; i <= d->model()->rowCount(); i++){
            rowNames.append(QString::number(i));
        }
    }

    nameCol.addItems(columnNames);
    finalCol.addItems(columnNames);
    startCol.addItems(columnNames);
    stopCol.addItems(columnNames);

    headerRow.addItems(rowNames);
    maxRow.addItems(rowNames);

    QPushButton okButton(tr("Ok"));
    QPushButton cancelButton(tr("Cancel"));

    QGridLayout l;

    QLabel nameLabel(tr("Name in Column:"));
    QLabel finalGradeLabel(tr("Final grade in column:"));
    QLabel startLabel(tr("Include Column:"));
    QLabel stopLabel(tr("up to (incl):"));
    QLabel headerLabel(tr("Grade title(s) in row:"));
    QLabel maxLabel(tr("Maximum points in row:"));

    l.addWidget(&nameLabel, 0, 0);
    l.addWidget(&finalGradeLabel, 1, 0);
    l.addWidget(&startLabel, 2, 0);
    l.addWidget(&stopLabel, 3, 0);
    l.addWidget(&headerLabel, 4, 0);
    l.addWidget(&maxLabel, 5, 0);

    l.addWidget(&nameCol, 0, 1);
    l.addWidget(&finalCol, 1, 1);
    l.addWidget(&startCol, 2, 1);
    l.addWidget(&stopCol, 3, 1);
    l.addWidget(&headerRow, 4, 1);
    l.addWidget(&maxRow, 5, 1);

    l.addWidget(&cancelButton, 6, 0);
    l.addWidget(&okButton, 6, 1);

    QDialog *dialog = new QDialog(this);

    dialog->setWindowTitle(tr("Generate Text:"));
    dialog->setLayout(&l);

    connect(&cancelButton, SIGNAL(clicked()), dialog, SLOT(reject()));
    connect(&okButton, SIGNAL(clicked()), dialog, SLOT(accept()));

    if(dialog->exec() == QDialog::Accepted){

        bool header = !(headerRow.currentText().startsWith(tr("<")));
        bool maxpoints = !(maxRow.currentText().startsWith(tr("<")));

        QString txt = tr("Beste ");

        if(nameCol.currentText().startsWith(tr("<"))){
            txt += tr("Student,\n\n");
        }
        else{
            txt += tr("#") + nameCol.currentText() + tr("#,\n\n");
        }

        if(!(finalCol.currentText().startsWith(tr("<")))){

            txt += tr("Je cijfer ");

            if(!m_courseCode->text().isEmpty()){
                txt += tr("voor het vak ") + m_courseCode->text() + tr(" ");
            }

            txt += tr("is een #") + finalCol.currentText() + tr("#");

            if(header){
                txt += tr("/#") + finalCol.currentText() + maxRow.currentText() + tr("#");
            }

            txt += tr(".\n\n");
        }

        if(!(startCol.currentText().startsWith(tr("<"))) && !(stopCol.currentText().startsWith(tr("<"))) && startCol.currentIndex() <= stopCol.currentIndex()){
            if(header){
                txt += tr("Specificatie (punten/totaal):\n");
            }
            else{
                txt += tr("Specificatie:\n");
            }

            for(int i = startCol.currentIndex(); i <= stopCol.currentIndex(); i++){
                if(header){
                    txt += tr("#") + startCol.itemText(i) + headerRow.currentText() + tr("#: ");
                }
                txt += tr("#") + startCol.itemText(i) + tr("#");
                if(maxpoints){
                    txt += tr("/#") + startCol.itemText(i) + maxRow.currentText() + tr("#");
                }
                txt += tr("\n");
            }
        }

        txt += tr("\n");
        txt += tr("Met vriendelijke groet,\n\n") + m_senderName->text() + tr("\n");

        m_edit->setText(txt);
    }
}

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
