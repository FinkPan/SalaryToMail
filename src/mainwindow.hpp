#ifndef MAINWINDOW_HPP
#define MAINWINDOW_HPP

#include <QMainWindow>

namespace Ui
{
class MainWindow;
}

class Data;
class ReadExcelFile;

class MainWindow: public QMainWindow
{
  Q_OBJECT

public:
  explicit MainWindow(QWidget *parent = 0);
  ~MainWindow();

private slots:
	void Open();
	void SetupData();

private:
  void CreateActions();
  void CreateTableData();


private:
  Ui::MainWindow *ui;
  QString file_name_;
  Data *data_;
  ReadExcelFile *ref_;



};

#endif // MAINWINDOW_H


