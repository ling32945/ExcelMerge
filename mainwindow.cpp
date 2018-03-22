#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFile>
#include <QFileDialog>
#include <QSettings>

#include <QMessageBox>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    QString filePath;
    readSettings("File", "ExcelDir", filePath);

    this->ui->lineEdit->setText(filePath);
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

void MainWindow::on_mergePushButton_clicked()
{
    for (int i = 0; i < ui->listWidget->count(); i++) {
        QListWidgetItem *listWidgetItem = ui->listWidget->item(i);
        QVariant itemData = listWidgetItem->data(Qt::UserRole);
        //QString excelFile = itemData.value<QString>();
        QString excelFile = itemData.toString();
        printf("Excel File: %s\n", excelFile.toStdString().c_str());
    }
}
