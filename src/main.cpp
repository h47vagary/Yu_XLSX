#include <iostream>
#include <string>
#include <vector>
#include <windows.h>    

#include "OpenXLSX.hpp"

using namespace OpenXLSX;

int main() {
    SetConsoleOutputCP(CP_UTF8); // 设置控制台输出为 UTF-8
    XLDocument doc;
    doc.open("E:/Code/Yu_XLSX/doc/揭阳2.xlsx");
    auto wb = doc.workbook();
    
    // 获取工作表数量
    unsigned int sheet_count = wb.worksheetCount();
    std::cout << "工作表总数: " << sheet_count << std::endl;

    // 遍历获取工作表名称
    auto sheetNames = wb.worksheetNames();
    for (size_t i = 0; i < sheetNames.size(); ++i) {
        std::cout << "Sheet " << i+1 << ": " << sheetNames[i] << std::endl;
    }

    doc.close();
    

    return 0;
}