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

    // 遍历每个要搜索的值，将结果按顺序存入 all_find_results
    std::vector<std::vector<CellPos>> all_find_results;
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

    // "消费时间" "车牌号码" "数量"
    std::string data_tag = "消费时间";
    std::string car_tag = "车牌号码";
    std::string num_tag = "数量";

    std::vector<CellPos> data_ = xlsx_find.find_cell_by_value(sheets_find[0], data_tag);
    std::vector<CellPos> car_ = xlsx_find.find_cell_by_value(sheets_find[0], car_tag);
    std::vector<CellPos> num_ = xlsx_find.find_cell_by_value(sheets_find[0], num_tag);

    // 获取列号，若未找到则设为 -1
    int data_col = data_.empty() ? -1 : data_[0].col;
    int car_col = car_.empty() ? -1 : car_[0].col;
    int num_col = num_.empty() ? -1 : num_[0].col;
    std::cout << "消费时间 列号: " << data_col << std::endl;
    std::cout << "车牌号码 列号: " << car_col << std::endl;
    std::cout << "数量 列号: " << num_col << std::endl;

    // 找出搜索内容在数据源中的 "消费时间", "车牌号码", "数量" 列的对应值
    struct find_data
    {
        std::string data;
        std::string car;
        std::string num;
    };
    std::vector<std::vector<find_data>> final_datas;

    for (size_t i = 0; i < all_find_results.size(); ++i)
    {
        std::vector<find_data> one_data;
        for (const auto& pos : all_find_results[i])
        {
            find_data fd;
            if (data_col != -1)
            {
                fd.data = xlsx_find.read_cell(sheets_find[0], pos.row, data_col);
            }
            if (car_col != -1)
            {
                fd.car = xlsx_find.read_cell(sheets_find[0], pos.row, car_col);
            }
            if (num_col != -1)
            {
                fd.num = xlsx_find.read_cell(sheets_find[0], pos.row, num_col);
            }
            one_data.push_back(fd);
        }
        final_datas.push_back(one_data);
    }

    // 输出结果
    std::cout << "\n最终结果: " << std::endl;
    for (size_t i = 0; i < range_data.size(); ++i)
    {
        std::cout << "值 '" << range_data[i] << "' 对应数据: " << std::endl;
        if (final_datas[i].empty())
        {
            std::cout << "未找到对应数据" << std::endl;
        }
        else
        {
            for (const auto& fd : final_datas[i])
            {
                std::cout << "消费时间: " << fd.data << ", 车牌号码: " << fd.car << ", 数量: " << fd.num << std::endl;
            }
        }
    }
    xlsx_find.close_file();


    XlsxHandle xlsx_result;
    // 创建新文件
    if (!xlsx_result.create_file("../doc/find_result.xlsx")) {
        std::cerr << "创建结果文件失败!" << std::endl;
        return 1;
    }

    // 新建“结果”工作表
    std::string result_sheet = "结果";
    if (!xlsx_result.create_sheet(result_sheet)) {
        std::cerr << "添加结果工作表失败!" << std::endl;
        return 1;
    }

    // 删除默认的 Sheet1
    auto result_sheets = xlsx_result.get_sheet_names();
    for (const auto& s : result_sheets) {
        if (s == "Sheet1") {
            xlsx_result.delete_sheet("Sheet1");
            break;
        }
    }

    // 写表头
    xlsx_result.write_cell(result_sheet, 1, 1, "日期");
    xlsx_result.write_cell(result_sheet, 1, 2, "车牌号码");
    xlsx_result.write_cell(result_sheet, 1, 3, "数量");

    // 写入数据
    unsigned int current_row = 2;
    for (size_t i = 0; i < range_data.size(); ++i)
    {
        if (!final_datas[i].empty())
        {
            for (const auto& fd : final_datas[i])
            {
                xlsx_result.write_cell(result_sheet, current_row, 1, fd.data);
                xlsx_result.write_cell(result_sheet, current_row, 2, fd.car);
                xlsx_result.write_cell(result_sheet, current_row, 3, fd.num);
                current_row++;
            }
        }
        //current_row++;
    }

    // 保存并关闭文件
    xlsx_result.save_file("../doc/find_result.xlsx");
    xlsx_result.close_file();


    return 0;
}