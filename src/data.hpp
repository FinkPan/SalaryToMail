#ifndef DATA_HPP
#define DATA_HPP

#include <QObject>
#include <QString>
#include <QVector>

class Data: public QObject
{
  Q_OBJECT


public:
  Data();
  ~Data();

signals:
  void WorkSeetNameChanged(QString);

public:
  QString GetValue(int n, int i);
  int cols();
  int rows();

  void Clear();

  void AddItems(QVector<QString> vstr);
  void SetWorkSheetName(QString name);

private:
  int item_num_;
  QString work_sheet_name_;
  QVector<QVector<QString>> items_;

};

#endif // !DATA_HPP
