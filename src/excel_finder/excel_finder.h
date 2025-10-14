/**
 * @file excel_finder.h
 * @brief Excel 查找业务逻辑
 * @version 1.0
 * @date 2025-10-12
 * 
 * @copyright Copyright (c) 2025
 * 
 * @todo 做成读取配置文件设置搜索条件的方式
 *          1. 文件获取方式：
 *                - 指定搜索目录下全部 xlsx 文件
 *                - 传入指定文件
 *          2. 搜索条件：
 *                - 指定 sheet 名称，或全部 sheet
 *                - 指定搜索范围
 *          3. 匹配项:
 *                - 是否区分大小写
 *                - 读取匹配的项，如"消费时间","车牌号码","数量" ...
 */

#pragma once

#include <vector>
#include <string>
#include "xlsx_handle.h"
#include "config_file.h"

/**
 * @brief 存储单行匹配到的最终数据
 */
struct FoundRecord {
    std::string data_time;   // 消费时间
    std::string car_number;  // 车牌号码
    std::string quantity;    // 数量
};

/**
 * 存储一个源值对应的所有匹配记录
 */
struct ValueResult {
    std::string source_value;               // 源值（如 "冀A84MA挂"）
    std::vector<FoundRecord> found_records; // 匹配到的记录
};

struct SheetResult {
    std::string sheet_name;                 // 源工作表名称
    std::string group_name;                 // 名称
    std::vector<ValueResult> value_results; // 每个源值对应的所有匹配记录
    std::vector<std::string> find_values;   // 该表中所有需要查找的值
};

/**
 * @brief Excel 查找业务处理类
 */
class ExcelFinder
{
public:
    ExcelFinder(const std::string& source_file_path,
                const std::string& target_file_path);
    ~ExcelFinder();

    // 打开文件
    bool init();

    // 设置查找列的标签
    void set_tags(const std::string& data_tag,
                  const std::string& car_tag,
                  const std::string& num_tag);

    // 设置从源文件读取数据的范围
    void set_source_read_range(unsigned int start_row,
                               unsigned int start_col,
                               unsigned int end_row,
                               unsigned int end_col);

    // 执行查找和结果数据提取
    bool execute();

    // 将查找结果导出到新的Excel文件
    bool export_results(const std::string& output_file_path);

    // 在控制台打印查找结果
    void print_results() const;

    // 重置结果表（即删除表后重新创建表）
    
    

private:
    XlsxHandle source_xlsx_;
    XlsxHandle target_xlsx_;
    std::string source_file_path = ""; 
    std::string target_file_path = ""; 

    std::string data_tag_;
    std::string car_tag_;
    std::string num_tag_;

    unsigned int read_start_row_ = 2;
    unsigned int read_start_col_ = 1;   // A
    unsigned int read_end_row_ = 9;
    unsigned int read_end_col_ = 7;     // G

    std::vector<SheetResult> results_;

    bool find_and_extract_data_from_target(const std::string& target_sheet_name,
                                           const std::vector<std::string>& source_value,
                                           ValueResult& out_value_result);

    void date_simplify(std::string& date_str);

};

