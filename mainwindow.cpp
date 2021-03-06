﻿#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFile>
#include <QFileDialog>
#include <QSettings>
#include <QRegExp>
#include <QProcess>

#include <QDebug>
#include <QMessageBox>

#include <iostream>

#ifdef _WIN32
  #include <windows.h>
#endif

#include "libxl.h"

using namespace libxl;

//QRegExp rxNum("^(-?\\d+)(\\.\\d+)?%?$");
QRegExp rxNum("^(-?\\d+)(\\.\\d+)?$");

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



QStringList MainWindow::generateDateList(QDate startDate, QDate endDate)
{
    QStringList dateHeaderList;

    QDate tempDate = startDate;
    tempDate.setDate(tempDate.year(), tempDate.month(), 1);

    int startMonth = startDate.month();
    if (startMonth % 3 != 0) {
        int remainder = startMonth % 3;
        tempDate = tempDate.addMonths(3 - remainder);
    }

    QString dateFormat = "yyyy-MM-dd";
    dateHeaderList.clear();

    int endDateYear = endDate.year();
    int endDateMonth = endDate.month();

    int year = 0, month = 0, day = 0;
    tempDate.getDate(&year, &month, &day);

    while(year < endDateYear || (year == endDateYear && month <= endDateMonth)) {
        tempDate.setDate(tempDate.year(), tempDate.month(), tempDate.daysInMonth());
        dateHeaderList.append(tempDate.toString(dateFormat));

        tempDate = tempDate.addMonths(3);
        tempDate.getDate(&year, &month, &day);
    }

    return dateHeaderList;
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
    QString fileName = QFileDialog::getSaveFileName(this, dialogTitle, oldFilePath,  QString::fromLocal8Bit("Excel 工作簿 (*.xls *.xlsx)"));
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

    Book* resultBook = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files

    resultBook->setKey(L"jae", L"windows-202220060ec5e50766b8646dacy5z7h0");


    Sheet* EPSSheet                         = resultBook->addSheet(L"基本每股收益");
    Sheet* netProfitSheet                   = resultBook->addSheet(L"净利润");
    Sheet* netProfitGrowthRateSheet         = resultBook->addSheet(L"净利润同比增长率");
    Sheet* nonNetProfitSheet                = resultBook->addSheet(L"扣非净利润");
    Sheet* nonNetProfitGrowthRateSheet      = resultBook->addSheet(L"扣非净利润同比增长率");
    Sheet* grossRevenueSheet                = resultBook->addSheet(L"营业总收入");
    Sheet* grossRevenueGrowthRateSheet      = resultBook->addSheet(L"营业总收入同比增长率");
    Sheet* BPSSheet                        = resultBook->addSheet(L"每股净资产");
    Sheet* ROESheet                         = resultBook->addSheet(L"净资产收益率");
    Sheet* ROEDitionSheet                   = resultBook->addSheet(L"净资产收益率-摊薄");
    Sheet* debtAssetRatioSheet              = resultBook->addSheet(L"资产负债率");
    Sheet* shareCapitalReserveSheet         = resultBook->addSheet(L"每股资本公积金");
    Sheet* retainedEarningsPerShareSheet    = resultBook->addSheet(L"每股未分配利润");
    Sheet* OCFPSSheet                       = resultBook->addSheet(L"每股经营现金流");
    Sheet* grossProfitRatioSheet            = resultBook->addSheet(L"销售毛利率");
    Sheet* depositTurnoverRatioSheet        = resultBook->addSheet(L"存款周转率");
    Sheet* netProfitMarginSheet             = resultBook->addSheet(L"销售净利润");

    EPSSheet->writeStr(1, 0, L"股票\\时间");
    netProfitSheet->writeStr(1, 0, L"股票\\时间");
    netProfitGrowthRateSheet->writeStr(1, 0, L"股票\\时间");
    nonNetProfitSheet->writeStr(1, 0, L"股票\\时间");
    nonNetProfitGrowthRateSheet->writeStr(1, 0, L"股票\\时间");
    grossRevenueSheet->writeStr(1, 0, L"股票\\时间");
    grossRevenueGrowthRateSheet->writeStr(1, 0, L"股票\\时间");
    BPSSheet->writeStr(1, 0, L"股票\\时间");
    ROESheet->writeStr(1, 0, L"股票\\时间");
    ROEDitionSheet->writeStr(1, 0, L"股票\\时间");
    debtAssetRatioSheet->writeStr(1, 0, L"股票\\时间");
    shareCapitalReserveSheet->writeStr(1, 0, L"股票\\时间");
    retainedEarningsPerShareSheet->writeStr(1, 0, L"股票\\时间");
    OCFPSSheet->writeStr(1, 0, L"股票\\时间");
    grossProfitRatioSheet->writeStr(1, 0, L"股票\\时间");
    depositTurnoverRatioSheet->writeStr(1, 0, L"股票\\时间");
    netProfitMarginSheet->writeStr(1, 0, L"股票\\时间");

    QDate startDate(1990, 1, 1);
    //QDate endDate = QDate::currentDate();
    QDate endDate(2017, 10, 1);
    QStringList dateList = generateDateList(startDate, endDate);

    QMap<QString, int> colIndexMap;

    for (int i = dateList.size() - 1; i >= 0; --i) {
        EPSSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        netProfitSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        netProfitGrowthRateSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        nonNetProfitSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        nonNetProfitGrowthRateSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        grossRevenueSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        grossRevenueGrowthRateSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        BPSSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        ROESheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        ROEDitionSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        debtAssetRatioSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        shareCapitalReserveSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        retainedEarningsPerShareSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        OCFPSSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        grossProfitRatioSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        depositTurnoverRatioSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());
        netProfitMarginSheet->writeStr(1, dateList.size() - i, dateList.at(i).toStdWString().c_str());

        colIndexMap[dateList.at(i)] = dateList.size() - i;
        //qDebug() << dateList.at(i) << " - " << colIndexMap[dateList.at(i)];
    }


    for (int i = 0; i < excelFileList.size(); ++i) {
//        if (i > 0)
//            break;

        QFileInfo excelFileInfo(excelFileList.at(i));
        if (!excelFileInfo.isReadable()) {
            QMessageBox::warning(this, QString::fromLocal8Bit("文件不可读"), QString::fromLocal8Bit("Excel 文件\n") + excelFileList.at(i) + QString::fromLocal8Bit("\n不可读！"));
            continue;
        }

        QString stockCode = excelFileInfo.baseName().left(6);
        printf("Stock Code: %s\n", stockCode.toStdString().c_str());

        // Read Excel
        Book* sourceBook = xlCreateBook();
        sourceBook->setKey(L"jae", L"windows-202220060ec5e50766b8646dacy5z7h0");
        if (sourceBook->load(excelFileList.at(i).toStdWString().c_str()) ) {
            Sheet* sourceSheet = sourceBook->getSheet(0);
            if (sourceSheet) {

                QMap<int, QString> dateColIndexMap;

                for (int row = sourceSheet->firstRow(); row < sourceSheet->lastRow(); ++row) {
                    //qDebug() << "Row No. " << row;
                    std::wcout << L"Row No. " << row << std::endl;

                    Sheet* operatorSheet = NULL;
                    int curRow = 0;
                    if (row < 2) {
                        if (row == 0) {
                            continue;
                        }
                    }
                    else {
                        if (row == 2) {
                            operatorSheet = EPSSheet;
                        }
                        else if (row == 3) {
                            operatorSheet = netProfitSheet;
                        }
                        else if (row == 4) {
                            operatorSheet = netProfitGrowthRateSheet;
                        }
                        else if (row == 5) {
                            operatorSheet = nonNetProfitSheet;
                        }
                        else if (row == 6) {
                            operatorSheet = nonNetProfitGrowthRateSheet;
                        }
                        else if (row == 7) {
                            operatorSheet = grossRevenueSheet;
                        }
                        else if (row == 8) {
                            operatorSheet = grossRevenueGrowthRateSheet;
                        }
                        else if (row == 9) {
                            operatorSheet = BPSSheet;
                        }
                        else if (row == 10) {
                            operatorSheet = ROESheet;
                        }
                        else if (row == 11) {
                            operatorSheet = ROEDitionSheet;
                        }
                        else if (row == 12) {
                            operatorSheet = debtAssetRatioSheet;
                        }
                        else if (row == 13) {
                            operatorSheet = shareCapitalReserveSheet;
                        }
                        else if (row == 14) {
                            operatorSheet = retainedEarningsPerShareSheet;
                        }
                        else if (row == 15) {
                            operatorSheet = OCFPSSheet;
                        }
                        else if (row == 16) {
                            operatorSheet = grossProfitRatioSheet;
                        }
                        else if (row == 17) {
                            operatorSheet = depositTurnoverRatioSheet;
                        }
                        else if (row == 18) {
                            operatorSheet = netProfitMarginSheet;
                        }
                        else {
                            break;
                        }

                        curRow = operatorSheet->lastRow();
                        //qDebug() << "Current Row NO. " << curRow;

                        // write the stock code
                        //qDebug() << "Stock Code 2" << stockCode;
                        //std::wcout << L"Test" << std::endl;
                        std::wcout << L"stock code " << stockCode.toStdWString().c_str() << std::endl;
                        operatorSheet->writeStr(curRow, 0, stockCode.toStdWString().c_str());
                    }

                    for (int col = sourceSheet->firstCol(); col < sourceSheet->lastCol(); ++col) {
                        if (col == 0) {
                            continue;
                        }

                        CellType cellType = sourceSheet->cellType(row, col);
                        //std::wcout << "(" << row << ", " << col << ") = ";
                        QString content = NULL;
                        if (sourceSheet->isFormula(row, col)) {
                            const wchar_t* s = sourceSheet->readFormula(row, col);
                            //std::wcout << (s ? s : L"null") << " [formula]";
                        }
                        else {
                            switch(cellType) {
                                case CELLTYPE_EMPTY: std::wcout << "[empty]"; break;
                                case CELLTYPE_NUMBER:
                                {
                                    double d = sourceSheet->readNum(row, col);
                                    //std::wcout << d << " [number]";
                                    content = QString::number(d);
                                    break;
                                }
                                case CELLTYPE_STRING:
                                {
                                    const wchar_t* s = sourceSheet->readStr(row, col);
                                    //std::wcout << (s ? s : L"null") << " [string]";
                                    content = QString::fromStdWString(s);
                                    break;
                                }
                                case CELLTYPE_BOOLEAN:
                                {
                                    bool b = sourceSheet->readBool(row, col);
                                    //std::wcout << (b ? "true" : "false") << " [boolean]";
                                    break;
                                }
                                case CELLTYPE_BLANK: std::wcout << "[blank]"; break;
                                case CELLTYPE_ERROR: std::wcout << "[error]"; break;
                            }
                        }
                        //std::wcout << std::endl;

                        if (row == 1) {
                            dateColIndexMap[col] = content;
                        }
                        else {
                            int curCol = colIndexMap[ dateColIndexMap[col] ];
                            if (rxNum.exactMatch(content)) {
                                operatorSheet->writeNum(curRow, curCol, content.toDouble());
                            }
                            else {
                                operatorSheet->writeStr(curRow, curCol, content.toStdWString().c_str());
                            }
                        }
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


    resultBook->save(L"report.xls");

    resultBook->release();

#ifdef _WIN32

    ::ShellExecuteA(NULL, "open", "report.xls", NULL, NULL, SW_SHOW);

#elif __APPLE__

    QProcess::execute("open report.xls");

#else

    QProcess::execute("oocalc report.xls");

#endif

}
