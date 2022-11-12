#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFile>
#include <QDebug>
#include <QTextStream>

#include <xlsxdocument.h>
#include <xlsxworksheet.h>
#include <xlsxformat.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>

using namespace QXlsx;


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    QFile file("://Input.csv");
    QFile file_out(":/OutPut.csv");
    QFile file_swe2("/home/mohammad/Documents/swe2.csv");
    QFile file_swe2_output("/home/mohammad/Documents/swe2_output.csv");
    QTextStream stream_swe2(&file_swe2);
    QTextStream stream_Swe2_output(&file_swe2_output);
    QXlsx::Document excel_file;
    QXlsx::Document excel_file_output;

    Format italic;
    italic.setFontItalic(true);
    Format red;
    red.setFontColor(Qt::red); // in :: be che manast inja?
    Format plain;


    //Exceledit("LKA","/home/mohammad/BBA.xlsx",":/3_input_LKA.csv");
     Exceledit(LKA,"LKA","/home/mohammad/ABB.xlsx",":/InPut.csv");

       ui->label->setText("The Projekt is DONE");
        qDebug()<<"enum : "<<LKA;
            }

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::Exceledit (int page,QString Page_name, QString output_xlsx_path, QString input_csv_path) /*(int a, int b, int c, int d, int e, int f, int g, int h , int i, int j, int k, int l, int n, int o
                            ,int w, int x, int y, int z, int aa,int bb,int cc, int dd, int ee, int ff)*/
{
    /*Page_name : page name output       ; for them
     *output_xlsx_path : excel_file output address
     *csv_path : csv   file input Path (which is added in source)
     *linelist[4+4].remove("\""),plain); //Data Type
     *linelist[1+4].remove("\""),plain);//unit
     *linelist[3+4].remove("\""),plain);//max value
     *linelist[2+4].remove("\""),plain);//min value
     *"PF_SteWhl_An","1","2","3","4","5 :rad unit",6: -23 min ,7: 23 max,"8: data type"," 9: cycle time 1ms","float32"
     */

    QFile file(input_csv_path);

    QXlsx::Document excel_file;

    Format italic;
    italic.setFontItalic(true);
    Format plain;

    excel_file.workbook()->setHtmlToRichStringEnabled(true);
    excel_file.addSheet(Page_name);

if (file.exists()){
qDebug()<<"the method file exist";
file.open(QFile::ReadOnly);
QTextStream stream(&file);
int line_number = 1;
int signal_number = 0 ;


   while(!stream.atEnd()){

        QString line;
        QStringList linelist;
        line = stream.readLine();
        linelist=line.split(",");
        signal_number++;
//        if (input_csv_path.contains("input",Qt::CaseInsensitive))
//            qDebug()<<"input :"<<linelist[9].remove("\"");
//        else
//            qDebug()<<"output :"<<linelist[9].remove("\"");


        if(linelist[page].remove("\"")=="x"){///////////////////////////////////////////////////////////////#########################################################
{

        qDebug()<<"input  : "<<linelist[0].remove("\"")<<" next coloum : "<<linelist[page].remove("\"");


        //qDebug()<<"method"<<line_number;


    RichString signal;
    QString str;
    str.setNum(signal_number);
    signal.addFragment( str,plain);
    //excel_file.write(line_number,1,str);

    //line_number++;
    RichString cell_format0;
    cell_format0.addFragment(Page_name, italic);
    if (input_csv_path.contains("input",Qt::CaseInsensitive))
        cell_format0.addFragment(" component shall receive the input signal ",plain); //
    else
        cell_format0.addFragment(" component shall send the output signal ",plain); //
    cell_format0.addFragment(linelist[0].remove("\""), italic);
    excel_file.write(line_number,1,cell_format0);
    line_number++;

    RichString cell_format1;
    cell_format1.addFragment("the signal ",plain);
    cell_format1.addFragment(linelist[0].remove("\""), italic);
    cell_format1.addFragment(" shall have the Data Type ",plain);
     if (input_csv_path.contains("input",Qt::CaseInsensitive))
    cell_format1.addFragment(linelist[10].remove("\""),plain); //Data Type
     else
    cell_format1.addFragment(linelist[9].remove("\""),plain); //Data Type
    excel_file.write(line_number,1,cell_format1);
    line_number++;
    RichString cell_format2;
    cell_format2.addFragment("the signal ",plain);
    cell_format2.addFragment(linelist[0].remove("\""), italic);
    cell_format2.addFragment(" shall have the unit ",plain);
    cell_format2.addFragment(linelist[5].remove("\""),plain);//unit
    excel_file.write(line_number,1,cell_format2);
    line_number++;
    RichString cell_format3;
    cell_format3.addFragment("the signal ",plain);
    cell_format3.addFragment(linelist[0].remove("\""), italic);
    cell_format3.addFragment(" shall have the resolution 0.001",plain);
    excel_file.write(line_number,1,cell_format3);
    line_number++;
    RichString cell_format4;
    cell_format4.addFragment("the signal ",plain);
    cell_format4.addFragment(linelist[0].remove("\""), italic);
    cell_format4.addFragment(" shall have the max value ",plain);
    cell_format4.addFragment(linelist[8].remove("\""),plain);//max value
    excel_file.write(line_number,1,cell_format4);
    line_number++;
    RichString cell_format5;
    cell_format5.addFragment("the signal ",plain);
    cell_format5.addFragment(linelist[0].remove("\""), italic);
    cell_format5.addFragment(" shall have the min value ",plain);
    cell_format5.addFragment(linelist[7].remove("\""),plain);//min value
    excel_file.write(line_number,1,cell_format5);
    line_number++;
    RichString cell_format6;
    cell_format6.addFragment("the signal ",plain);
    cell_format6.addFragment(linelist[0].remove("\""), italic);
    cell_format6.addFragment(" shall have the default value XXX ",plain);
    excel_file.write(line_number,1,cell_format6);
    line_number++;
    if (input_csv_path.contains("input",Qt::CaseInsensitive))
    {
    RichString cell_format7;
    cell_format7.addFragment("the signal ",plain);
    cell_format7.addFragment(linelist[0].remove("\""), italic);
    cell_format7.addFragment(" shall have the cycle time ",plain);
    cell_format7.addFragment(linelist[9].remove("\""),plain); //cycle time
    excel_file.write(line_number,1,cell_format7);
    line_number++;
    }

    line_number++;
    excel_file.saveAs(output_xlsx_path);
}
}
}
}
}
