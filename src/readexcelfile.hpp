#ifndef READEXCELFILE_HPP
#define READEXCELFILE_HPP

#include <QString>
#include <QObject>

class Data;

class ReadExcelFile : public QObject
{
  Q_OBJECT

public:
  ReadExcelFile();
  ~ReadExcelFile();

signals:
  void setupdata();

public:
  int Open1(QString file_name,Data *work_sheet);

private:

};

#endif // !READEXCELFILE_HPP
