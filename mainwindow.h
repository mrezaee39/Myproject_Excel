#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QFile>
#include <QDebug>
#include <QTextStream>

#include <xlsxdocument.h>
#include <xlsxworksheet.h>
#include <xlsxformat.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>



QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

enum page {
    RG= 1,
    LKA,
    MCIG,
    SM,
    FOC,
    Platform,
};


class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    //void Exceledit(QString s5, QString s6, QString s7, QString s8);/*(, int a, int b, int c, int d, int e, int f, int g, int h , int i, int j, int k, int l, int n, int o
      //             ,int w, int x, int y, int z, int aa,int bb,int cc, int dd, int ee, int ff);*/

    void Exceledit(int page,QString Page_name, QString output_xlsx_path, QString input_csv_path);
    QXlsx::RichString  Cell_format(QString phrase1, QString signlaname, QString phrase3, QString phrase4);
    void Exceledit_with_function(int page_number, QString Page_name_, QString input_path, QString output_path);



private slots:


private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
