#include "readexcelfile.hpp"
#include "data.hpp"

#include <QAxObject>
#include <QDebug> 

ReadExcelFile::ReadExcelFile()
{
}

ReadExcelFile::~ReadExcelFile()
{
}

int ReadExcelFile::Open1(QString file_name,Data *work_sheet)
{
  QAxObject*  excel = nullptr;
  QAxObject*  workbooks = nullptr;
  QAxObject*  workbook = nullptr;

  excel = new QAxObject("Excel.Application");
  if (excel->isNull())
    return -1;

  excel->setProperty("Visible", false);

  workbooks = excel->querySubObject("Workbooks");
  if (workbooks == nullptr || workbooks->isNull())
    return -1;


  workbook = workbooks->querySubObject("Open(QString, QVariant,QVariant)", file_name, 0, "True");
  if (workbook == nullptr || workbook->isNull())
    return -1;

  QAxObject *active_sheet = workbook->querySubObject("ActiveSheet");

  work_sheet->Clear();
  work_sheet->SetWorkSheetName(active_sheet->property("Name").toString());

  QAxObject * usedrange = active_sheet->querySubObject("UsedRange");
  if (usedrange == nullptr || usedrange->isNull())
    return -1;

  QVariant var;
    var = usedrange->dynamicCall("Value");
  delete usedrange;

  QVariantList varRows = var.toList();
  if (varRows.isEmpty())
    return -1;

  QVariantList varCols;
  QVector<QString> vstr;

  for (int r = 1; r < varRows.size(); r++)
  {
    varCols = varRows[r].toList();
    vstr.clear();
    for (int c = 0; c <varCols.size(); c++)
    {
      vstr.push_back(varCols[c].toString());
    }
    work_sheet->AddItems(vstr);
  }

  workbooks->dynamicCall("Close()");
  excel->dynamicCall("Quit()");
  delete excel;

  emit setupdata();
  return 0;

}
