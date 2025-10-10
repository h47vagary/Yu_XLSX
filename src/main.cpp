#include <iostream>
#include <windows.h>    

#include "xlsx_handle.h"

int main() {
    SetConsoleOutputCP(CP_UTF8); // 设置控制台输出为 UTF-8

    // 打开搜索源文件
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

    unsigned int range_read_start_row = 2, range_read_start_col = 1, range_read_end_row = 9, range_read_end_col = 7;
    auto range_data = xlsx.read_range(sheets[0], range_read_start_row, range_read_start_col, range_read_end_row, range_read_end_col);
    std::cout << "读取区域 (" << range_read_start_row << "," << range_read_start_col << ") to (" << range_read_end_row << "," << range_read_end_col << ") 内容: " << std::endl;
    for (const auto &cell_val : range_data)
    {
        std::cout << cell_val << " | ";
    }

    xlsx.close_file();


    // 打开搜索目标文件
    XlsxHandle xlsx_find;
    if (!xlsx_find.open_file("../doc/福邦2-4交易详情135.xlsx"))
    {
        std::cerr << "文件打开失败!" << std::endl;
        return 1;
    }

    auto sheets_find = xlsx_find.get_sheet_names();
    std::cout << "工作表数量: " << sheets_find.size() << std::endl;
    for (const auto &name : sheets_find)
    {
        std::cout << "工作表: " << name << std::endl;
    }

    std::vector<std::vector<CellPos>> all_find_results;
    // 遍历每个要搜索的值，将结果按顺序存入 all_find_results
    for (const auto& range_data_one : range_data) // 注意加 const&，避免拷贝
    {
        auto find_results = xlsx_find.find_cell_by_value(sheets_find[0], range_data_one);
        all_find_results.push_back(find_results); // 正确：结果顺序与 range_data 完全对应
    }

    // 输出结果（此时 i 对应的就是 range_data[i] 的查找结果）
    std::cout << "\n查找结果: " << std::endl;
    for (size_t i = 0; i < range_data.size(); ++i)
    {
        std::cout << "值 '" << range_data[i] << "' 出现位置: ";
        if (all_find_results[i].empty())
        {
            std::cout << "未找到";
        }
        else
        {
            for (const auto& pos : all_find_results[i])
            {
                // 可选：将列号转为字母（如 1→A），更直观
                std::cout << "(" << pos.row << "," << pos.col << ") ";
            }
        }
        std::cout << std::endl;
    }
    xlsx_find.close_file();



    
    return 0;
}