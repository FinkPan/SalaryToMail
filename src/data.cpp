#include "data.hpp"

#include <QDebug> 

Data::Data()
{
}

Data::~Data()
{
}


QString Data::GetValue(int n, int i)
{
  Q_ASSERT(n < items_.size() && i < items_[n].size());
  return items_[n][i];
}

int Data::cols()
{
  Q_ASSERT(!items_.empty());
  return items_[0].size();
}

int Data::rows()
{
  return items_.size();
}

void Data::Clear()
{
  items_.clear();
}

void Data::AddItems(QVector<QString> vstr)
{
  items_.push_back(vstr);
}

void Data::SetWorkSheetName(QString name)
{
  work_sheet_name_ = name;
  emit WorkSeetNameChanged(work_sheet_name_);
}
