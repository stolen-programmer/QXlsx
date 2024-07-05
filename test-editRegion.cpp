//
// Created by 20264 on 2024/5/13.
//

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"

#include <QDebug>
int main()
{
    QXlsx::Document d{R"(C:\Users\20264\Desktop\工作簿1.xlsx)"};

    qDebug() << d.load();

    auto &r = reinterpret_cast<QXlsx::Worksheet *>(d.workbook()->activeSheet())->editRegion();

    for (auto &cellR : r) {
        qDebug() << cellR.firstRow() << cellR.firstColumn() << cellR.lastRow()
                 << cellR.lastColumn();
    }
    return 0;
}