#include <QRegExp>
#include <QDateTime>
#include <QDebug>
#include <QElapsedTimer>
#include <QFileDialog>
#include "mainwidget.h"
#include "ui_mainwidget.h"

MainWidget::MainWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::MainWidget)
{
    ui->setupUi(this);
    m_xls = new ExcelBase;
}

MainWidget::~MainWidget()
{
    delete ui;
}
#define MY_DEBUG
void MainWidget::on_pushButtonOpen_clicked()
{
    QString xlsFile = QFileDialog::getOpenFileName(this,QString(),QString(),"excel(*.xls *.xlsx)");
    debugText(xlsFile);
    if(xlsFile.isEmpty())
        return;

    m_xls->open(xlsFile);

    debugText("sheetNames "+m_xls->sheetNames().join(";"));
    m_xls->setCurrentSheet(1);
    m_datas.clear();
    m_xls->readAll(m_datas);//QList< QList<QVariant> >

    debugText(QString("读取完毕：%1列，%2行").arg(m_datas.at(0).size()).arg(m_datas.size()));

//    foreach (QList<QVariant> OneClumn, m_datas) {//
//        foreach (QVariant v, OneClumn) {
//            qDebug()<<v.toString()<<"\t";
//        }
//        qDebug()<<"\r\n";
//    }


    QList< QList<QVariant> > newDatas;
    QList< QList<QVariant> > saveDatas;
    QMap<QString,int> numInDate5;
    QMap<QString,int> numInDate10;
    QMap<QString,int> numInDate0;
    QMap<QString,int> numInDate_;
    QList<QString> dateList;
    int ALL5=0,ALL10=0,ALL0=0,ALL_=0;

    enum ZhangDanType{Xin,Jiu};
    ZhangDanType type;
    if(m_datas.size()!=0){
        if(m_datas.at(0).at(0).toString()=="消费明细报表"){
            type = Jiu;
        }else if(m_datas.at(0).at(0).toString()=="消费时间"){
            type = Xin;
        }
        else
            type = Xin;
        int ShiJianIndex = 0;
        int NameIndex = 2;
        int ManeyIndex = (type==Xin)?6:4;
        int StartRow = (type==Xin)?1:2;
        for(int i=StartRow;i<m_datas.size();i++){
            QList<QVariant> oneRow;
            oneRow<<m_datas.at(i).at(ShiJianIndex)
                  <<m_datas.at(i).at(NameIndex)
                  <<m_datas.at(i).at(ManeyIndex);
            newDatas<<oneRow;

            QString d = m_datas.at(i).at(ShiJianIndex).toString();
            d.resize(10);
            if(dateList.indexOf(d)==-1)
                dateList.append(d);
            int maney = m_datas.at(i).at(ManeyIndex).toInt();
            switch(maney){
            case 5:
                numInDate5[d] += 1;
                ALL5++;
                break;
            case 10:
                numInDate10[d] += 1;
                ALL10++;
                break;
            case 0:
                numInDate0[d] += 1;
                ALL0++;
                break;
            default:
                numInDate_[d] += 1;
                ALL_++;
                break;
            }
        }
        debugText("提取完毕："+(type==Xin)?"新":"旧");
        qDebug()<<numInDate5;
        qDebug()<<numInDate10;
        qDebug()<<numInDate0;
        qDebug()<<numInDate_;
        debugText(QString("数量 %1").arg(newDatas.size()));
        debugText("第一 "+newDatas.at(0).at(0).toString()+"  "+newDatas.at(0).at(1).toString()+"  "+newDatas.at(0).at(2).toString());
        debugText("最后 "+newDatas.at(newDatas.size()-1).at(0).toString()+"  "+newDatas.at(newDatas.size()-1).at(1).toString()+"  "+newDatas.at(newDatas.size()-1).at(2).toString());

        debugText(QString("日期\t5\t10\t0\t*"));
        foreach (QString s, dateList) {
            debugText(QString("'%1\t%2\t%3\t%4\t%5").arg(s)
                      .arg(numInDate5[s])
                      .arg(numInDate10[s])
                      .arg(numInDate0[s])
                      .arg(numInDate_[s])
                      );
        }
        debugText(QString("合计\t%1\t%2\t%3\t%4")
                  .arg(ALL5)
                  .arg(ALL10)
                  .arg(ALL0)
                  .arg(ALL_)
                  );

        int pageRow = 60;
        int pageClumn = 3;
        int size = newDatas.size();
        QList<QVariant> title;
        title<<"时间"<<"姓名"<<"￥"<<" "<<"时间"<<"姓名"<<"￥"<<" "<<"时间"<<"姓名"<<"￥";
        for(int i=0;i<size;){

            if(i%(pageRow*pageClumn)==0){
                qDebug()<<i<<" title----------------";
                saveDatas<<title;
            }

            QList<QVariant> oneRow;

            oneRow<<"'"+newDatas.at(i).at(0).toString()<<newDatas.at(i).at(1)<<newDatas.at(i).at(2)<<" ";
            int ii = i+pageRow;
            int iii = ii+pageRow;
            if(ii<size){
                oneRow<<"'"+newDatas.at(ii).at(0).toString()<<newDatas.at(ii).at(1)<<newDatas.at(ii).at(2)<<" ";
                if(iii<size)
                    oneRow<<"'"+newDatas.at(iii).at(0).toString()<<newDatas.at(iii).at(1)<<newDatas.at(iii).at(2);
            }
            //qDebug()<<oneRow;
            saveDatas<<oneRow;

            i++;
            if(i%pageRow==0)
                i+=pageRow*pageClumn-pageRow;
            qDebug()<<i;
        }
    }else{
        qDebug()<<"read error!";
    }
    bool b;
    b=m_xls->setCurrentSheet(2);
    debugText(QString("setCurrentSheet %1").arg(b));
    b=m_xls->writeCurrentSheet(saveDatas);
    debugText(QString("writeCurrentSheet %1").arg(b));

    QRegExp exp("(.*/)(.*)(\\.xls)");
    int pos = exp.indexIn(xlsFile);
    QString path,filename,fileFormat,savePathName;
    QString dateTime = QDateTime::currentDateTime().toString("yyMMdd-hhmmss");
    if(pos>-1&&exp.captureCount()==3){
        path = exp.cap(1);
        filename = exp.cap(2);
        fileFormat = exp.cap(3);
        filename += "--"+dateTime;
        savePathName = path+filename+fileFormat;
    }else
        savePathName = dateTime+".xls";
    m_xls->saveAs(savePathName);
    debugText("另存为 "+savePathName);

    m_xls->close();
}

void MainWidget::on_pushButton_clicked()
{
//    QList< QList<QVariant> > saveDatas;
//    QList<QVariant> title;
//    title<<"时间"<<"姓名"<<"￥"<<" "<<"时间"<<"姓名"<<"￥"<<" "<<"时间"<<"姓名"<<"￥";
//    saveDatas<<title;

//    ExcelBase saveXls;
//    qDebug()<<"create "<<saveXls.create("1.xls");
//    qDebug()<<"setCurrentSheet "<<saveXls.setCurrentSheet(1);
//    qDebug()<<"writeCurrentSheet "<<saveXls.writeCurrentSheet(saveDatas);
//    saveXls.save();
//    return;

    ExcelBase *m_xls;
    m_xls = new ExcelBase;
    QString xlsFile = QFileDialog::getExistingDirectory(this);
    if(xlsFile.isEmpty())
        return;
    xlsFile += "/excelRWByCztr1988.xls";
    QElapsedTimer timer;
    timer.start();
    m_xls->create(xlsFile);
    qDebug()<<"create cost:"<<timer.elapsed()<<"ms";timer.restart();
    QList< QList<QVariant> > m_datas;
    for(int i=0;i<100;++i)
    {
        QList<QVariant> rows;
        for(int j=0;j<100;++j)
        {
            rows.append(i*j);
        }
        m_datas.append(rows);
    }
    m_xls->setCurrentSheet(1);
    timer.restart();
    m_xls->writeCurrentSheet(m_datas);
    qDebug()<<"write cost:"<<timer.elapsed()<<"ms";timer.restart();
    m_xls->save();
}

void MainWidget::debugText(const QString &s)
{
    ui->plainTextEdit->appendPlainText(s);
}
