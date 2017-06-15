#include "mainwindow.hpp"
#include "ui_mainwindow.h"
#include "data.hpp"
#include "readexcelfile.hpp"

#include <QObject>
#include <QAxObject>
#include <QMessageBox>
#include <QFileDialog>
#include <QHeaderView>
#include <QDebug> 


MainWindow::MainWindow(QWidget *parent):
QMainWindow(parent),
    ui(new Ui::MainWindow)
{
  data_ = new(Data);
  ref_ = new(ReadExcelFile);

  ui->setupUi(this);

  CreateTableData();

  connect(ui->actionopen, &QAction::triggered, this, &MainWindow::Open);
  connect(data_, &Data::WorkSeetNameChanged, ui->label, &QLabel::setText);
  connect(ref_, &ReadExcelFile::setupdata, this, &MainWindow::SetupData);

}

MainWindow::~MainWindow()
{
  delete ui;
}

void MainWindow::Open() 
{
  file_name_ = QFileDialog::getOpenFileName(
    this,
    tr("Open Excel File"),
    "",
    tr("Excel File (*.xlsx *.xls)"));

  if (file_name_.isEmpty())
    return;

  ref_->Open1(file_name_, data_);

}

void MainWindow::CreateTableData()
{
  ui->tableWidget->horizontalHeader()->hide();
  ui->tableWidget->verticalHeader()->hide();

  ui->tableWidget->setColumnCount(16);
  ui->tableWidget->setRowCount(2);
  ui->tableWidget->setSpan(0, 0, 2, 1);
  ui->tableWidget->setSpan(0, 1, 2, 1);
  ui->tableWidget->setSpan(0, 2, 2, 1);
  ui->tableWidget->setSpan(0, 3, 1, 2);
  ui->tableWidget->setSpan(0, 5, 2, 1);
  ui->tableWidget->setSpan(0, 6, 2, 1);
  ui->tableWidget->setSpan(0, 7, 2, 1);
  ui->tableWidget->setSpan(0, 8, 2, 1);
  ui->tableWidget->setSpan(0, 9, 1, 4);
  ui->tableWidget->setSpan(0, 13, 2, 1);
  ui->tableWidget->setSpan(0, 14, 2, 1);
  ui->tableWidget->setSpan(0, 15, 2, 1);

  QTableWidgetItem *item1 = new QTableWidgetItem(QString::fromLocal8Bit("工号"));
  item1->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item2 = new QTableWidgetItem(QString::fromLocal8Bit("姓名"));
  item2->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item3 = new QTableWidgetItem(QString::fromLocal8Bit("基本工资"));
  item3->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item4 = new QTableWidgetItem(QString::fromLocal8Bit("工龄工资"));
  item4->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item5 = new QTableWidgetItem(QString::fromLocal8Bit("工龄补助"));
  item5->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item6 = new QTableWidgetItem(QString::fromLocal8Bit("公司工龄"));
  item6->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item7 = new QTableWidgetItem(QString::fromLocal8Bit("浮动"));
  item7->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item8 = new QTableWidgetItem(QString::fromLocal8Bit("提成"));
  item8->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item9 = new QTableWidgetItem(QString::fromLocal8Bit("本月绩效"));
  item9->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item10 = new QTableWidgetItem(QString::fromLocal8Bit("高温补贴"));
  item10->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item11 = new QTableWidgetItem(QString::fromLocal8Bit("代扣款"));
  item11->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item12 = new QTableWidgetItem(QString::fromLocal8Bit("养老保险"));
  item12->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item13 = new QTableWidgetItem(QString::fromLocal8Bit("医疗保险"));
  item13->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item14 = new QTableWidgetItem(QString::fromLocal8Bit("失业保险"));
  item14->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item15 = new QTableWidgetItem(QString::fromLocal8Bit("个税"));
  item15->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item16 = new QTableWidgetItem(QString::fromLocal8Bit("应发工资"));
  item16->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item17 = new QTableWidgetItem(QString::fromLocal8Bit("实际工资"));
  item17->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);
  QTableWidgetItem *item18 = new QTableWidgetItem(QString::fromLocal8Bit("邮箱"));
  item18->setTextAlignment(Qt::AlignHCenter| Qt::AlignVCenter);

  ui->tableWidget->setItem(0, 0, item1);
  ui->tableWidget->setItem(0, 1, item2);
  ui->tableWidget->setItem(0, 2, item3);
  ui->tableWidget->setItem(0, 3, item4);
  ui->tableWidget->setItem(1, 3, item5);
  ui->tableWidget->setItem(1, 4, item6);
  ui->tableWidget->setItem(0, 5, item7);
  ui->tableWidget->setItem(0, 6, item8);
  ui->tableWidget->setItem(0, 7, item9);
  ui->tableWidget->setItem(0, 8, item10);
  ui->tableWidget->setItem(0, 9, item11);
  ui->tableWidget->setItem(1, 9, item12);
  ui->tableWidget->setItem(1, 10, item13);
  ui->tableWidget->setItem(1, 11, item14);
  ui->tableWidget->setItem(1, 12, item15);
  ui->tableWidget->setItem(0, 13, item16);
  ui->tableWidget->setItem(0, 14, item17);
  ui->tableWidget->setItem(0, 15, item18);
  ui->tableWidget->resizeColumnsToContents();

}

void MainWindow::SetupData()
{
  ui->tableWidget->setRowCount(data_->rows()+2);

  for (int r = 0; r < data_->rows(); r++)
  {
    for (int c = 0; c < data_->cols(); c++)
    {
      ui->tableWidget->setItem(r+2, c, new QTableWidgetItem(data_->GetValue(r, c)));
    }
  }
  ui->tableWidget->resizeColumnsToContents();
}
