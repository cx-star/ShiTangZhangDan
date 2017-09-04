#ifndef MAINWIDGET_H
#define MAINWIDGET_H

#include <QWidget>
#include "ExcelBase.h"

namespace Ui {
class MainWidget;
}

class MainWidget : public QWidget
{
    Q_OBJECT

public:
    explicit MainWidget(QWidget *parent = 0);
    ~MainWidget();

private slots:
    void on_pushButtonOpen_clicked();

    void on_pushButton_clicked();

private:
    Ui::MainWidget *ui;

    ExcelBase *m_xls;
    QList< QList<QVariant> > m_datas;
    void myDebugText(const QString& s);
};

#endif // MAINWIDGET_H
