#include "excel_finder.h"
#include <iostream>
#include <filesystem>

#define XLSX_FIND_RESULT_FONT_SIZE 11

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
    std::cout << "[ExcelFinder] 源文件包含 " << source_sheets.size() << " 个工作表." << std::endl;
    for (auto source_sheet : source_sheets) {
        std::cout << "  - " << source_sheet << std::endl;
    }   

    // 获取目标文件的第一个工作表
    auto target_sheets = target_xlsx_.get_sheet_names();
    if (target_sheets.empty()) {
        std::cerr << "[ExcelFinder] 错误: 目标文件没有工作表." << std::endl;
        return false;
    }
    const std::string &target_sheet = target_sheets[0];

    // 预缓存目标表数据，构建索引
    target_xlsx_.cache_sheet_data(target_sheet);
    target_xlsx_.build_index(target_sheet, false);

    // 查找目标表列索引
    ColumnIndex col_idx;
    auto find_col = [&](const std::string& tag)->int {
        auto pos = target_xlsx_.find_cell_by_value_fast(target_sheet, tag, false);
        return pos.empty()? -1 : pos[0].col;
    };
    col_idx.data_col = find_col(data_tag_);
    col_idx.car_col = find_col(car_tag_);
    col_idx.num_col = find_col(num_tag_);

    if (col_idx.car_col == -1)
    {
        std::cerr << "[ExcelFinder]错误: 未找到车牌列" <<  car_tag_ << "." << std::endl;
        return false;
    }
    if (col_idx.data_col == -1) {
        std::cerr << "[ExcelFinder]错误：未找到日期列" << data_tag_ << "." << std::endl;
    }
    if (col_idx.num_col == -1) {
        std::cerr << "[ExcelFinder]错误：未找到数量列" << num_tag_ << "." << std::endl;
    }

    // 遍历所有源工作表
    for (const auto& sheet_name : source_sheets) {
        SheetResult sheet_result;
        sheet_result.sheet_name = sheet_name;
        // 读取名称
        sheet_result.group_name = source_xlsx_.read_cell(sheet_name, 1, 2); // 读取B1单元格
        // 从源工作表读取数据
        auto source_values = source_xlsx_.read_range(sheet_name, read_start_row_, read_start_col_, read_end_row_, read_end_col_, false);
        if (source_values.empty()) {
            std::cout << "[ExcelFinder] 警告: 源工作表 " << sheet_name << " 没有数据." << std::endl;
            results_.push_back(sheet_result);
            continue;
        }
        // 保留源工作表待搜索数据
        sheet_result.find_values = source_values;

        std::cout << "[ExcelFinder] 处理源工作表: " << sheet_name 
                  << " (名称: " << sheet_result.group_name << ", 待搜索项数: " << source_values.size() << ")" << std::endl;

        // 遍历每个源值，进行查找
        for (const auto& source_value : source_values) {
            ValueResult value_result;
            value_result.source_value = source_value;

            // 单个 source_value 调用
            find_and_extract_data_from_target_fast(target_sheet, source_value, value_result, col_idx);
            
            if (!value_result.found_records.empty()) {
                sheet_result.value_results.push_back(value_result);
            }
                
        }

        results_.push_back(sheet_result);
    }

    std::cout << "[ExcelFinder] 查找执行完成." << std::endl;
    return true;
}

bool ExcelFinder::find_and_extract_data_from_target_fast(const std::string& target_sheet_name,
                                                         const std::string& source_value,
                                                         ValueResult& out_value_result,
                                                         const ColumnIndex& col_idx)
{
    auto positions = target_xlsx_.find_cell_by_value_fast(target_sheet_name, source_value, false);
    if (positions.empty()) {
        std::cout << "can't find value: " << source_value << " int sheet name: " << target_sheet_name << std::endl;
        return false;
    }

    const auto& table = target_xlsx_.get_sheet_cache(target_sheet_name);

    for (const auto& pos : positions) {
        FoundRecord record;
        if (col_idx.data_col != -1) record.data_time  = table[pos.row-1][col_idx.data_col-1];
        if (col_idx.car_col  != -1) record.car_number = table[pos.row-1][col_idx.car_col-1];
        if (col_idx.num_col  != -1) record.quantity   = table[pos.row-1][col_idx.num_col-1];

        std::cout << "[sheet_name]: " << target_sheet_name << " found " << "[data_time]: " << record.data_time 
                    << " [car_number]: " << record.car_number << " [quantity]: " << record.quantity << std::endl;
        date_simplify(record.data_time);
        out_value_result.found_records.push_back(record);
    }

    return true;
}

void ExcelFinder::date_simplify(std::string &date_str)
{
    if (date_str.size() >= 10) {
        date_str = date_str.substr(0, 10);
    }
}


void ExcelFinder::print_results() const
{
    std::cout << "\n==================== 查找结果汇总 ====================" << std::endl;
    for (const auto& sheet_result : results_) {
        std::cout << "--- 源工作表: " << sheet_result.sheet_name << " --- 名称: " << sheet_result.group_name << std::endl;

        if (sheet_result.value_results.empty()) {
            std::cout << "  未找到任何匹配项。" << std::endl;
            continue;
        }

        for (const auto& value_result : sheet_result.value_results) {
            std::cout << "  值: '" << value_result.source_value << "' 找到 " << value_result.found_records.size() << " 条记录:" << std::endl;
            for (const auto& record : value_result.found_records) {
                std::cout << "    日期: " << record.data_time
                          << ", 车牌: " << record.car_number
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
    xlsx_result.write_cell(result_sheet, 1, 1, "名称");
    xlsx_result.write_cell(result_sheet, 1, 2, "搜索源");
    xlsx_result.write_cell(result_sheet, 1, 3, "日期");
    xlsx_result.write_cell(result_sheet, 1, 4, "车牌号码");
    xlsx_result.write_cell(result_sheet, 1, 5, "数量");
    for (int col = 1; col <= 5; ++col) {
        xlsx_result.set_font_size(result_sheet, 1, col, XLSX_FIND_RESULT_FONT_SIZE);
    }

    // 6. 写入数据
    unsigned int current_row = 2;
    for (const auto& sheet_result : results_) {

        bool first_entry_in_sheet = true; // 标记是否为该表的第一条记录

        // 收集整个块内的所有记录
        std::vector<FoundRecord> block_records;
        for (const auto& value_result : sheet_result.value_results) {
            for (auto record : value_result.found_records) {
                date_simplify(record.data_time);            // 日期简化
                block_records.push_back(record);
            }
        }

        // 对整个块按日期排序
        std::sort(block_records.begin(), block_records.end(),
                  [](const FoundRecord& a, const FoundRecord& b) {
                      return a.data_time < b.data_time;
                  });

        // 查看搜索源文件的size和block_records的size
        size_t record_count = block_records.size();
        size_t source_count = sheet_result.find_values.size();
        size_t max_count = std::max(record_count, source_count);

        for (size_t i = 0; i < max_count; ++i) {
            // 名称列
            if (i == 0) {
                xlsx_result.write_cell(result_sheet, current_row, 1, sheet_result.group_name);
                xlsx_result.set_font_size(result_sheet, current_row, 1, XLSX_FIND_RESULT_FONT_SIZE);
            }

            // 搜索源
            if (i < source_count)
                xlsx_result.write_cell(result_sheet, current_row, 2, sheet_result.find_values[i]);
                xlsx_result.set_font_size(result_sheet, current_row, 2, XLSX_FIND_RESULT_FONT_SIZE);

            // 日期、车牌、数量
            if (i < record_count) {
                xlsx_result.write_cell(result_sheet, current_row, 3, block_records[i].data_time);
                xlsx_result.write_cell(result_sheet, current_row, 4, block_records[i].car_number);
                xlsx_result.write_cell(result_sheet, current_row, 5, block_records[i].quantity);
                xlsx_result.set_font_size(result_sheet, current_row, 3, XLSX_FIND_RESULT_FONT_SIZE);
                xlsx_result.set_font_size(result_sheet, current_row, 4, XLSX_FIND_RESULT_FONT_SIZE);
                xlsx_result.set_font_size(result_sheet, current_row, 5, XLSX_FIND_RESULT_FONT_SIZE);
            }
            current_row++;
        }

        // 空行分隔不同工作表的结果
        current_row += 2;
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