#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#ifndef _UNICODE
#define _UNICODE
#endif

#include <QMainWindow>

#include <QDate>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_browsePushButton_clicked();

    void on_loadPushButton_clicked();

    void on_destBrowsePushButton_clicked();

    void on_mergePushButton_clicked();

private:
    Ui::MainWindow *ui;

    void writeSettings(QString group, QString key, QString value);

    void readSettings(QString group, QString key, QString &value);

    QStringList generateDateList(QDate startDate, QDate endDate);
};

#endif // MAINWINDOW_H
