#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFile>
#include <QFileDialog>
#include <QSettings>
#include <QProcess>

#include <QDebug>
#include <QMessageBox>

#include <iostream>

#ifdef _WIN32
  #include <windows.h>
#endif

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
    if (this->ui->destFilePathLineEdit->text().isEmpty()) {
        QMessageBox::critical(this, QString::fromLocal8Bit("保存"), QString::fromLocal8Bit("请选择要保存的结果文件的路径！") );
    }

    QStringList excelFileList;
    for (int i = 0; i < ui->listWidget->count(); i++) {
        QListWidgetItem *listWidgetItem = ui->listWidget->item(i);
        if (listWidgetItem->checkState() == Qt::Unchecked) {
            continue;
        }
        QVariant itemData = listWidgetItem->data(Qt::UserRole);
        //QString excelFile = itemData.value<QString>();
        QString excelFile = itemData.toString();
        excelFileList.append(excelFile);
        printf("Excel File: %s\n", excelFile.toStdString().c_str());
    }

    if (excelFileList.size() <= 0) {
        QMessageBox::warning(this, QString::fromLocal8Bit("合并"), QString::fromLocal8Bit("没有可合并的文件，请选择需要合并的文件！"));
        return;
    }

    QString testC;
    Book* resultBook = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files
    Sheet* EPSSheet = resultBook->addSheet(L"基本每股收益");
    Sheet* netProfitSheet = resultBook->addSheet(L"净利润");
    Sheet* netProfitGrowthRateSheet = resultBook->addSheet(L"净利润同比增长率");
    Sheet* nonNetProfitSheet = resultBook->addSheet(L"扣非净利润");
    Sheet* nonNetProfitGrowthRateSheet = resultBook->addSheet(L"扣非净利润同比增长率");

    for (int i = 0; i < excelFileList.size(); ++i) {
        QFileInfo excelFileInfo(excelFileList.at(i));
        if (!excelFileInfo.isReadable()) {
            QMessageBox::warning(this, QString::fromLocal8Bit("文件不可读"), QString::fromLocal8Bit("Excel 文件\n") + excelFileList.at(i) + QString::fromLocal8Bit("\n不可读！"));
            continue;
        }

        QString stockCode = excelFileInfo.baseName().left(6);

        // Read Excel
        Book* sourceBook = xlCreateBook();
        if (sourceBook->load(excelFileList.at(0).toStdWString().c_str()) ) {
            Sheet* sourceSheet = sourceBook->getSheet(0);
            if (sourceSheet) {
                // write the stock code
                EPSSheet->writeStr(2, 0, stockCode.toStdWString().c_str());
                for (int row = sourceSheet->firstRow(); row < sourceSheet->lastRow(); ++row) {
                    for (int col = sourceSheet->firstCol(); col < sourceSheet->lastCol(); ++col) {
                        CellType cellType = sourceSheet->cellType(row, col);
                        std::wcout << "(" << row << ", " << col << ") = ";
                        if (sourceSheet->isFormula(row, col)) {
                            const wchar_t* s = sourceSheet->readFormula(row, col);
                            std::wcout << (s ? s : L"null") << " [formula]";
//                        const char* s = sourceSheet->readFormula(row, col);
//                        std::cout << (s ? s : "null") << " [formula]";
                        }
                        else {
                            switch(cellType) {
                                case CELLTYPE_EMPTY: std::wcout << "[empty]"; break;
                                case CELLTYPE_NUMBER:
                                {
                                    double d = sourceSheet->readNum(row, col);
                                    std::wcout << d << " [number]";
                                    break;
                                }
                                case CELLTYPE_STRING:
                                {
                                    const wchar_t* s = sourceSheet->readStr(row, col);
                                    std::wcout << (s ? s : L"null") << " [string]";
//                                const char* s = sourceSheet->readStr(row, col);
                                    if (row == 1 && col == 0) {
                                        testC = QString::fromStdWString(s);
                                        //QString msg = QString(QLatin1String(s)).toLatin1();
                                        qDebug() << "Row: " << row << " - Col: " << row << " --- " << testC;
                                        //qDebug() << "Row: " << row << " - Col: " << row << " --- " << msg;
                                    }
                                    break;
                                }
                                case CELLTYPE_BOOLEAN:
                                {
                                    bool b = sourceSheet->readBool(row, col);
                                    std::wcout << (b ? "true" : "false") << " [boolean]";
//                                    std::cout << (b ? "true" : "false") << " [boolean]";
                                    break;
                                }
                                case CELLTYPE_BLANK: std::wcout << "[blank]"; break;
                                case CELLTYPE_ERROR: std::wcout << "[error]"; break;
                            }
                        }
                        std::wcout << std::endl;
                    }
                }
            }
            else {
                const char* sheetError = sourceBook->errorMessage();
                qDebug() << "Error!!! " << sheetError;
            }
        }
        else {
            const char* errorMsg = sourceBook->errorMessage();
            qDebug() << "Error!!! " << errorMsg;
        }
        sourceBook->release();
    }


//    Sheet* sheet = book->addSheet(L"基本每股收益");
//
//    printf("TestC %s", testC.toStdString().c_str());
//
//    sheet->writeStr(2, 1, L"Hello, World !");
//    sheet->writeStr(2, 2, testC.toStdWString().c_str());
//
//    sheet->writeStr(3,1, L"中文");
//    sheet->writeNum(4, 1, 1000);
//    sheet->writeNum(5, 1, 2000);
//
//    Font* font = book->addFont();
//    font->setColor(COLOR_RED);
//    font->setBold(true);
//    Format* boldFormat = book->addFormat();
//    boldFormat->setFont(font);
//    sheet->writeFormula(6, 1, L"SUM(B5:B6)", boldFormat);
//
//    Format* dateFormat = book->addFormat();
//    dateFormat->setNumFormat(NUMFORMAT_DATE);
//    sheet->writeNum(8, 1, book->datePack(2011, 7, 20), dateFormat);
//
//    sheet->setCol(1, 1, 12);

    resultBook->save(L"report.xls");

    resultBook->release();

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
