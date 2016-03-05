#include "mainwindow.h"

#include <QWidget>
#include <QTableView>
#include <QTableWidget>
#include <QPushButton>
#include <QLabel>

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QGridLayout>

#include <QFileDialog>
#include <QInputDialog>
#include <QMessageBox>

#include <QRegExp>
#include <QStringRef>

#include <QtXlsx>

#include "xlsxsheetmodel.h"


//using namespace QXlsx;

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent)
{
    this->setWindowTitle(tr("HR ELE Emailgenerator"));

    m_smtp = NULL;

    QWidget *mainWidget = new QWidget(this);
    mainWidget->setMinimumSize(1000, 800);

    QVBoxLayout *l1 = new QVBoxLayout();

    l1->addWidget(createGeneralOptionsWidget());
    l1->addWidget(createMiddleWidget());
    l1->addWidget(createXlsxViewer());

    mainWidget->setLayout(l1);

    this->setCentralWidget(mainWidget);

    updateText();

    //qDebug() << xlsx.cellAt(1,1)->value() <<  endl;
}

QWidget* MainWindow::createRowSelectWidget(){

    QWidget *rowSelectWidget = new QWidget(this);
    rowSelectWidget->setMinimumWidth(150);
    rowSelectWidget->setMaximumWidth(150);
    rowSelectWidget->setMinimumHeight(380);
    rowSelectWidget->setMaximumHeight(380);

    QVBoxLayout *rowSelectLayout = new QVBoxLayout();
    rowSelectLayout->setAlignment(Qt::AlignTop);

    m_start = new QComboBox(this);
    m_stop = new QComboBox(this);
    m_preview = new QComboBox(this);
    connect(m_start, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));
    connect(m_stop, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));
    connect(m_preview, SIGNAL(currentTextChanged(QString)), this, SLOT(updateInfo()));

    m_nMails = new QLCDNumber(4, this);
    m_nMails->setSegmentStyle(QLCDNumber::Filled);
    m_nMails->setPalette(QPalette(Qt::red));

    m_emailColumn = new QComboBox(this);
    connect(m_emailColumn, SIGNAL(currentTextChanged(QString)), this, SLOT(updateText()));
    m_emailAppend = new QLineEdit(tr("@hr.nl"), this);
    connect(m_emailAppend, SIGNAL(textChanged(QString)), this, SLOT(updateText()));

    rowSelectLayout->addWidget(new QLabel(tr("Preview Email#:"), this));
    rowSelectLayout->addWidget(m_preview);
    rowSelectLayout->addWidget(new QLabel(tr("Start from row:"), this));
    rowSelectLayout->addWidget(m_start);
    rowSelectLayout->addWidget(new QLabel(tr("to row (included):"), this));
    rowSelectLayout->addWidget(m_stop);
    rowSelectLayout->addWidget(new QLabel(tr("# of Mails:"), this));
    rowSelectLayout->addWidget(m_nMails);
    rowSelectLayout->addWidget(new QLabel(tr("Email address\nis in Column:"), this));
    rowSelectLayout->addWidget(m_emailColumn);
    rowSelectLayout->addWidget(new QLabel(tr("and append to this:"), this));
    rowSelectLayout->addWidget(m_emailAppend);

    rowSelectWidget->setLayout(rowSelectLayout);

    return rowSelectWidget;
}

QWidget* MainWindow::createGeneralOptionsWidget(){
    QWidget *generalOptionsWidget = new QWidget(this);

    QGridLayout *generalOptionsLayout = new QGridLayout();
    generalOptionsLayout->setAlignment(Qt::AlignTop);

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

    QPushButton *SMTPConnectButton = new QPushButton(tr("Connect to mailserver"), this);
    connect(SMTPConnectButton, SIGNAL(clicked()), this, SLOT(SMTPconnect()));

    QPushButton *sendMailsButton = new QPushButton(tr("Send mails now"), this);
    connect(sendMailsButton, SIGNAL(clicked()), this, SLOT(sendMails()));

    generalOptionsLayout->addWidget(new QLabel(tr("Sender Name:"), this), 0, 0);
    generalOptionsLayout->addWidget(m_senderName, 0, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Sender Email:"), this), 1, 0);
    generalOptionsLayout->addWidget(m_senderEmail, 1, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Add Bcc:"), this), 2, 0);
    generalOptionsLayout->addWidget(m_emailBcc, 2, 1);
    generalOptionsLayout->addWidget(new QLabel(tr("Email Subject:"), this), 0, 3);
    generalOptionsLayout->addWidget(m_emailSubject, 0, 4);
    generalOptionsLayout->addWidget(new QLabel(tr("Course Code:"), this), 1, 3);
    generalOptionsLayout->addWidget(m_courseCode, 1, 4);
    generalOptionsLayout->addWidget(SMTPConnectButton, 0, 5);
    generalOptionsLayout->addWidget(sendMailsButton, 2, 5);

    generalOptionsWidget->setLayout(generalOptionsLayout);

    return generalOptionsWidget;
}

QWidget* MainWindow::createEditorWidget(){

    QWidget *editorWidget = new QWidget(this);
    editorWidget->setMinimumHeight(380);

    QVBoxLayout *l = new QVBoxLayout();
    l->setAlignment(Qt::AlignTop);

    m_edit = new QTextEdit(this);
    m_edit->setText(tr("Beste #B#,\n\nJe studentnummer is #A#, de vakcode is #B1#.\n\nJe cijfer is #C#.\n\nMet vriendelijke groet,\n\n#B2#\n"));
    connect(m_edit, SIGNAL(textChanged()), this, SLOT(updateText()));

    l->addWidget(new QLabel(tr("Edit:")));
    l->addWidget(m_edit);

    editorWidget->setLayout(l);

    return editorWidget;
}

QWidget* MainWindow::createPreviewWidget(){

    QWidget *previewWidget = new QWidget(this);
    previewWidget->setMinimumHeight(380);

    QVBoxLayout *l = new QVBoxLayout();

    m_view = new QTextEdit(this);
    m_view->setReadOnly(true);

    l->addWidget(new QLabel(tr("Preview:")));
    l->addWidget(m_view);

    previewWidget->setLayout(l);

    return previewWidget;
}

QWidget* MainWindow::createMiddleWidget(){

    QWidget *middleWidget = new QWidget();
    middleWidget->setMinimumHeight(400);

    QHBoxLayout *l = new QHBoxLayout();

    l->addWidget(createEditorWidget());
    l->addWidget(createRowSelectWidget());
    l->addWidget(createPreviewWidget());

    middleWidget->setLayout(l);

    return middleWidget;
}

QWidget* MainWindow::createXlsxViewer(){

    QWidget *xlsxWidget = new QWidget(this);
    xlsxWidget->setMinimumHeight(200);
    xlsxWidget->setMinimumWidth(600);

    QVBoxLayout *l = new QVBoxLayout();
    QHBoxLayout *l1 = new QHBoxLayout();

    QPushButton *loadXlsxButton = new QPushButton(tr("Load xlsx file"), this);
    connect(loadXlsxButton, SIGNAL(clicked()), this, SLOT(loadSheet()));

    m_tabWidget = new QTabWidget(this);
    m_tabWidget->setTabPosition(QTabWidget::South);
    connect(m_tabWidget, SIGNAL(currentChanged(int)), this, SLOT(updateSheet()));

    l1->addWidget(new QLabel(tr("XLSX Viewer:")));
    l1->addWidget(loadXlsxButton);
    l->addLayout(l1);
    l->addWidget(m_tabWidget);

    xlsxWidget->setLayout(l);

    return xlsxWidget;
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

void MainWindow::blockUpdate(bool b){
    m_start->blockSignals(b);
    m_stop->blockSignals(b);
    m_preview->blockSignals(b);
}

void MainWindow::updateSheet(){
    QTableView *d = (QTableView*)m_tabWidget->currentWidget();

    m_emailColumn->clear();

    QStringList headers;
    for(int i = 0; i < d->model()->columnCount(); i++)
    {
      headers.append(d->model()->headerData(i, Qt::Horizontal).toString());
    }

    m_emailColumn->addItems(headers);

    updateInfo();
}

void MainWindow::updateInfo(){
    blockUpdate(true);

    QTableView *d = (QTableView*)m_tabWidget->currentWidget();
    int max = d->model()->rowCount();
    int start = m_start->currentText().toInt();
    int stop = m_stop->currentText().toInt();
    int preview = m_preview->currentText().toInt();

    m_start->clear();
    m_stop->clear();
    m_preview->clear();

    for(int i = 0; i < max; i++){
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
    m_preview->setCurrentText(QString::number(preview));

    blockUpdate(false);

    m_nMails->display(m_stop->currentText().toInt() - m_start->currentText().toInt() + 1);

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

    m_smtp = new SmtpClient("smtp.hr.nl", 465, SmtpClient::SslConnection);

    QString user = QInputDialog::getText(this, tr("Username"), tr("SMTP Username:"), QLineEdit::Normal, m_senderEmail->text(), &ok);

    if(!ok){
        SMTPdisconnect();
        return;
    }

    m_smtp->setUser(user);

    QString password = QInputDialog::getText(this, tr("Password"), tr("SMTP Password for ") + user + tr(":"), QLineEdit::Password, tr(""), &ok);

    if(!ok){
        SMTPdisconnect();
        return;
    }

    m_smtp->setPassword(password);

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

void MainWindow::SMTPdisconnect(){
    m_smtp->quit();
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

        if(m_smtp == NULL){
            //QMessageBox::warning(this, tr("Error:"), tr("SMTP not connected! Sending mail to ") + recv_mail + tr(" failed!"));
            failed += tr("  ") + recv_mail + tr("\n");
            nFailed++;
            continue;
        }

        if(!m_smtp->sendMail(message)){
            //QMessageBox::warning(this, tr("Error:"), tr("Sending mail to ") + recv_mail + tr(" failed!"));
            failed += tr("  ") + recv_mail + tr("\n");
            nFailed++;
            continue;
        }

        success += tr("  ") + recv_mail + tr("\n");
        nSuccess++;

    }

    QMessageBox::information(this, tr("Info:"), tr("Tried to send ") + nMails + tr("mails.\n\n") + tr("Mails OK: ") + nSuccess + tr("\n\nMails Failed:") + nFailed + tr("\n") + failed + tr("\n"));
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

        /*for(int i = col.length(); i < 0; i++){
            char chr = col.data()[i].toLatin1();
            c +=  qPow(26, col.length() - i - 1) * (chr - 'A' + 1);
        }*/

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

MainWindow::~MainWindow()
{

}
