#include <iostream>
#include <windows.h>    

#include "xlsx_handle.h"

int main() {
    SetConsoleOutputCP(CP_UTF8); // 设置控制台输出为 UTF-8

    XlsxHandle xlsx;

    if (!xlsx.open_file("../doc/福邦1.xlsx"))
    {
        std::cerr << "文件打开失败!" << std::endl;
        return 1;
    }

    auto sheets = xlsx.get_sheet_names();
    std::cout << "工作表数量: " << sheets.size() << std::endl;
    for (const auto &name : sheets)
    {
        std::cout << "工作表: " << name << std::endl;
    }

    if (!sheets.empty())
    {
        std::string val = xlsx.read_cell(sheets[0], 1, 1);
        std::cout << "第一个工作表第(1,1)单元格内容: " << val << std::endl;

        std::string val2_4 = xlsx.read_cell(sheets[0], 2, 4);
        std::cout << "第一个工作表第(2,4)单元格内容: " << val2_4 << std::endl;

        std::string val2_5 = xlsx.read_cell(sheets[0], 2, 5);
        std::cout << "第一个工作表第(2,5)单元格内容: " << val2_5 << std::endl;
    }

    xlsx.close_file();

    return 0;
}