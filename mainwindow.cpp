#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFile>
#include <QFileDialog>
#include <QSettings>
#include <QProcess>

#include <QMessageBox>

#include "libxl.h"

using namespace libxl;

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    QString filePath;
    readSettings("File", "ExcelDir", filePath);
    if (!filePath.isEmpty()) {
        this->ui->lineEdit->setText(filePath);
    }
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::writeSettings(QString group, QString key, QString value)
{
    QSettings settings("Moose Soft Consulting", "ExcelMerge");

    settings.setValue(group + "/" + key, value);

}

void MainWindow::readSettings(QString group, QString key, QString &value)
{
    value.clear();
    QSettings settings("Moose Soft", "Clipper");

    settings.value(group + "/" + key, value);
}

void MainWindow::on_browsePushButton_clicked()
{
    QString oldFilePath = this->ui->lineEdit->text();
    if (oldFilePath.isEmpty()) {
        oldFilePath = QDir::homePath();
    }

    QString dialogTile = QString::fromLocal8Bit("选择文件夹");
    QString filePath = QFileDialog::getExistingDirectory(this, dialogTile, oldFilePath);
    ui->lineEdit->setText(filePath);
}

void MainWindow::on_loadPushButton_clicked()
{
    QString filePath = this->ui->lineEdit->text();
    writeSettings("File", "ExcelDir", filePath);
    QDir excelsDir(filePath);

    ui->listWidget->clear();

    if (!excelsDir.exists()) {
        QMessageBox::warning(this, QString::fromLocal8Bit("文件夹不存在"), QString::fromLocal8Bit("文件夹：\n%1\n不存在，请选择正确的Excel文件夹！").arg(filePath));
        return;
    }

    QStringList fileSuffixFilters;
    fileSuffixFilters << "*.xls" << "*.xlsx";

    excelsDir.setNameFilters(fileSuffixFilters);
    QStringList excelNameList = excelsDir.entryList();
    //this->ui->listView->setModel(excelNameList);
    for (int i = 0; i < excelNameList.size(); i++) {
        QListWidgetItem *listWidgetItem = new QListWidgetItem();
        listWidgetItem->setText(excelNameList.at(i));
        listWidgetItem->setCheckState(Qt::Checked);
        QString excelFile = filePath + "/" + excelNameList.at(i);
        QVariant itemData;
        itemData.setValue(excelFile);
        listWidgetItem->setData(Qt::UserRole, itemData);

        ui->listWidget->addItem(listWidgetItem);
    }
}

void MainWindow::on_destBrowsePushButton_clicked()
{
    QString oldFilePath = this->ui->destFilePathLineEdit->text();
    if (oldFilePath.isEmpty()) {
        oldFilePath = QDir::homePath();
    }

    QString dialogTitle = QString::fromLocal8Bit("保存");
    QString fileName = QFileDialog::getSaveFileName(this, dialogTitle, oldFilePath);
    this->ui->destFilePathLineEdit->setText(fileName);
}

void MainWindow::on_mergePushButton_clicked()
{
    for (int i = 0; i < ui->listWidget->count(); i++) {
        QListWidgetItem *listWidgetItem = ui->listWidget->item(i);
        QVariant itemData = listWidgetItem->data(Qt::UserRole);
        //QString excelFile = itemData.value<QString>();
        QString excelFile = itemData.toString();
        printf("Excel File: %s\n", excelFile.toStdString().c_str());
    }


    Book* book = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files

    Sheet* sheet = book->addSheet("Sheet1");

    sheet->writeStr(2, 1, "Hello, World !");
    sheet->writeNum(4, 1, 1000);
    sheet->writeNum(5, 1, 2000);

    Font* font = book->addFont();
    font->setColor(COLOR_RED);
    font->setBold(true);
    Format* boldFormat = book->addFormat();
    boldFormat->setFont(font);
    sheet->writeFormula(6, 1, "SUM(B5:B6)", boldFormat);

    Format* dateFormat = book->addFormat();
    dateFormat->setNumFormat(NUMFORMAT_DATE);
    sheet->writeNum(8, 1, book->datePack(2011, 7, 20), dateFormat);

    sheet->setCol(1, 1, 12);

    book->save("report.xls");

    book->release();

    //ui->pushButton->setText("Please wait...");
    //ui->pushButton->setEnabled(false);

#ifdef _WIN32

    ::ShellExecuteA(NULL, "open", "report.xls", NULL, NULL, SW_SHOW);

#elif __APPLE__

    QProcess::execute("open report.xls");

#else

    QProcess::execute("oocalc report.xls");

#endif

    //ui->pushButton->setText("Generate Excel Report");
    //ui->pushButton->setEnabled(true);
}
