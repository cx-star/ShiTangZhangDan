#include <QDebug>
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
//#define MY_DEBUG
void MainWidget::on_pushButtonOpen_clicked()
{
    QString xlsFile = QFileDialog::getOpenFileName(this,QString(),QString(),"excel(*.xls *.xlsx)");
    if(xlsFile.isEmpty())
        return;

    m_xls->open(xlsFile);

    qDebug()<<"names "<<m_xls->sheetNames();
    m_xls->setCurrentSheet(1);
    m_datas.clear();
    m_xls->readAll(m_datas);//QList< QList<QVariant> >
#ifdef MY_DEBUG
    qDebug()<<"read over";
    qDebug()<<"clumn "<<m_datas.at(0).size()<<" row "<<m_datas.size();;
    foreach (QList<QVariant> OneClumn, m_datas) {
        foreach (QVariant v, OneClumn) {
            qDebug()<<v.toString()<<"\t";
        }
        qDebug()<<"\r\n";
    }
#endif

    QList< QList<QVariant> > newDatas;
    QList< QList<QVariant> > saveDatas;

    enum ZhangDanType{Xin,Jiu};
    ZhangDanType type;
    if(m_datas.size()!=0){
        if(m_datas.at(0).at(0).toString()=="消费明细报表"){
            type = Jiu;
        }else if(m_datas.at(0).at(0).toString()=="消费时间"){
            type = Xin;
        }
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
        }
        qDebug()<<"size  "<<newDatas.size();
        qDebug()<<"first "<<newDatas.at(0);
        qDebug()<<"last  "<<newDatas.at(newDatas.size()-1);

        int pageRow = 50;
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

            oneRow<<newDatas.at(i).at(0)<<newDatas.at(i).at(1)<<newDatas.at(i).at(2)<<" ";
            int ii = i+pageRow;
            int iii = ii+pageRow;
            if(ii<size){
                oneRow<<newDatas.at(ii).at(0)<<newDatas.at(ii).at(1)<<newDatas.at(ii).at(2)<<" ";
                if(iii<size)
                    oneRow<<newDatas.at(iii).at(0)<<newDatas.at(iii).at(1)<<newDatas.at(iii).at(2);
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
    //qDebug()<<"addBook "<<m_xls->addBook();
    //qDebug()<<"setCurrentSheet "<<m_xls->setCurrentSheet(m_xls->sheetNames().size()-1);
    qDebug()<<"setCurrentSheet "<<m_xls->setCurrentSheet(2);
    qDebug()<<"writeCurrentSheet "<<m_xls->writeCurrentSheet(saveDatas);
    ExcelBase saveXls;
    saveXls.create("1.xls");
    saveXls.setCurrentSheet(1);
    saveXls.writeCurrentSheet(saveDatas);
    saveXls.save();

    m_xls->close();
}
