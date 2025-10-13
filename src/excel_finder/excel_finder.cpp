#include "excel_finder.h"
#include <iostream>
#include <filesystem>
#include <tuple>

ExcelFinder::ExcelFinder(const std::string &source_file_path, const std::string &target_file_path)
    : source_file_path(source_file_path),
      target_file_path(target_file_path)
{
}

ExcelFinder::~ExcelFinder()
{
    source_xlsx_.close_file();
    target_xlsx_.close_file();
}

bool ExcelFinder::init()
{
    std::cout << "[ExcelFinder] 正在初始化..." << std::endl;
    if (!source_xlsx_.open_file(source_file_path)) {
        std::cerr << "[ExcelFinder] 错误: 无法打开源文件: " << source_file_path << std::endl;
        return false;
    }
    if (!target_xlsx_.open_file(target_file_path)) {
        std::cerr << "[ExcelFinder] 错误: 无法打开目标文件: " << target_file_path << std::endl;
        source_xlsx_.close_file();
        return false;
    }
    std::cout << "[ExcelFinder] 初始化完成." << std::endl;
    return true;
}

void ExcelFinder::set_tags(const std::string &data_tag, const std::string &car_tag, const std::string &num_tag)
{
    data_tag_ = data_tag;
    car_tag_ = car_tag;
    num_tag_ = num_tag;
}

void ExcelFinder::set_source_read_range(unsigned int start_row,
                                        unsigned int start_col,
                                        unsigned int end_row,
                                        unsigned int end_col)
{
    read_start_row_ = start_row;
    read_start_col_ = start_col;
    read_end_row_ = end_row;
    read_end_col_ = end_col;
}

bool ExcelFinder::execute()
{
    std::cout << "[ExcelFinder] 开始执行查找..." << std::endl;
    results_.clear();

    // 获取源文件的所有工作表
    auto source_sheets = source_xlsx_.get_sheet_names();
    if (source_sheets.empty()) {
        std::cerr << "[ExcelFinder] 错误: 源文件没有工作表." << std::endl;
        return false;
    }

    // 获取目标文件的第一个工作表
    auto target_sheets = target_xlsx_.get_sheet_names();
    if (target_sheets.empty()) {
        std::cerr << "[ExcelFinder] 错误: 目标文件没有工作表." << std::endl;
        return false;
    }
    const std::string &target_sheet = target_sheets[0];

    // 遍历所有源工作表
    for (const auto& sheet_name : source_sheets) {
        SheetResult sheet_result;
        sheet_result.sheet_name = sheet_name;

        // 从源工作表读取数据
        auto source_values = source_xlsx_.read_range(sheet_name, read_start_row_, read_start_col_, read_end_row_, read_end_col_, false);
        if (source_values.empty()) {
            std::cout << "[ExcelFinder] 警告: 源工作表 " << sheet_name << " 没有数据." << std::endl;
            results_.push_back(sheet_result);
            continue;
        }

        // 遍历每个源值，进行查找
        for (const auto& source_value : source_values) {
            ValueResult value_result;
            value_result.source_value = source_value;

            if (find_and_extract_data_from_target(target_sheet, source_values, value_result)) {
                sheet_result.value_results.push_back(value_result);
            }
        }

        results_.push_back(sheet_result);
    }

    std::cout << "[ExcelFinder] 查找执行完成." << std::endl;
    return true;
}

bool ExcelFinder::find_and_extract_data_from_target(const std::string& target_sheet_name, const std::vector<std::string>& source_values, ValueResult& value_result) 
{
    // 在目标工作表中查找当前源值
    auto positions = target_xlsx_.find_cell_by_value(target_sheet_name, value_result.source_value);
    if (positions.empty()) {
        return false; // 未找到，无需处理
    }

    // 查找目标表中的列标题位置
    int data_col = -1, car_col = -1, num_col = -1;
    auto data_pos = target_xlsx_.find_cell_by_value(target_sheet_name, data_tag_);
    auto car_pos = target_xlsx_.find_cell_by_value(target_sheet_name, car_tag_);
    auto num_pos = target_xlsx_.find_cell_by_value(target_sheet_name, num_tag_);

    if (!data_pos.empty()) data_col = data_pos[0].col;
    if (!car_pos.empty()) car_col = car_pos[0].col;
    if (!num_pos.empty()) num_col = num_pos[0].col;

    // 如果关键列（如车牌号码）未找到，无法提取有效信息
    if (car_col == -1) {
        std::cerr << "[ExcelFinder] 警告: 在目标工作表中未找到列 '" << car_tag_ << "', 无法提取数据。" << std::endl;
        return false;
    }

    // 遍历所有匹配到的位置，提取整行数据
    for (const auto& pos : positions) {
        FoundRecord record;
        if (data_col != -1) {
            record.data_time = target_xlsx_.read_cell(target_sheet_name, pos.row, data_col);
        }
        if (car_col != -1) {
            record.car_number = target_xlsx_.read_cell(target_sheet_name, pos.row, car_col);
        }
        if (num_col != -1) {
            record.quantity = target_xlsx_.read_cell(target_sheet_name, pos.row, num_col);
        }
        value_result.found_records.push_back(record);
    }

    return true;
}

void ExcelFinder::print_results() const 
{
    std::cout << "\n==================== 查找结果汇总 ====================" << std::endl;
    for (const auto& sheet_result : results_) {
        std::cout << "--- 源工作表: " << sheet_result.sheet_name << " ---" << std::endl;
        if (sheet_result.value_results.empty()) {
            std::cout << "  该工作表在目标文件中未找到任何匹配项。" << std::endl;
            continue;
        }
        for (const auto& value_result : sheet_result.value_results) {
            std::cout << "  值: '" << value_result.source_value << "' 找到 " << value_result.found_records.size() << " 条记录:" << std::endl;
            for (const auto& record : value_result.found_records) {
                std::cout << "    消费时间: " << record.data_time 
                          << ", 车牌号码: " << record.car_number 
                          << ", 数量: " << record.quantity << std::endl;
            }
        }
    }
    std::cout << "======================================================" << std::endl;
}

bool ExcelFinder::export_results(const std::string& output_file_path_base) 
{
    std::cout << "[ExcelFinder] 正在导出结果到文件: " << output_file_path_base << "..." << std::endl;

    std::string real_file_path = output_file_path_base;
    XlsxHandle xlsx_result;

    // 1. 判断文件是否存在，若存在则直接打开
    if (std::filesystem::exists(real_file_path)) {
        if (!xlsx_result.open_file(real_file_path)) {
            std::cerr << "[ExcelFinder] 打开已存在文件失败!" << std::endl;
            return false;
        }
    } else {
        // 文件不存在则创建
        if (!xlsx_result.create_file(real_file_path)) {
            std::cerr << "[ExcelFinder] 创建结果文件失败!" << std::endl;
            return false;
        }
    }

    // 2. 自动生成结果表名（结果-1, 结果-2 …）
    int idx = 1;
    std::string result_sheet;
    do {
        result_sheet = "结果-" + std::to_string(idx++);
    } while (xlsx_result.is_sheet_exist(result_sheet));

    // 3. 创建结果表
    if (!xlsx_result.create_sheet(result_sheet)) {
        std::cerr << "[ExcelFinder] 添加结果工作表失败!" << std::endl;
        xlsx_result.close_file();
        return false;
    }

    // 4. 删除默认Sheet1（如果存在且不是结果表）
    if (xlsx_result.is_sheet_exist("Sheet1") && result_sheet != "Sheet1") {
        xlsx_result.delete_sheet("Sheet1");
    }

    // 5. 写表头
    xlsx_result.write_cell(result_sheet, 1, 1, "日期");
    xlsx_result.write_cell(result_sheet, 1, 2, "车牌号码");
    xlsx_result.write_cell(result_sheet, 1, 3, "数量");

    // 6. 写入数据
    unsigned int current_row = 2;
    for (const auto& sheet_result : results_) {
        for (const auto& value_result : sheet_result.value_results) {
            for (const auto& record : value_result.found_records) {
                xlsx_result.write_cell(result_sheet, current_row, 1, record.data_time);
                xlsx_result.write_cell(result_sheet, current_row, 2, record.car_number);
                xlsx_result.write_cell(result_sheet, current_row, 3, record.quantity);
                current_row++;
            }
        }
        // 不同源工作表结果之间留一行空白
        current_row++;
    }

    // 7. 保存文件
    if (!xlsx_result.save_file(real_file_path)) {
        std::cerr << "[ExcelFinder] 保存结果文件失败!" << std::endl;
        xlsx_result.close_file();
        return false;
    }

    xlsx_result.close_file();
    std::cout << "[ExcelFinder] 结果已成功导出到表 " << result_sheet << std::endl;
    return true;
}