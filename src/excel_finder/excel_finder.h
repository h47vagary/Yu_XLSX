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
#include <map>
#include <algorithm>
#include "xlsx_handle.h"


/**
 * @brief 存储单行匹配到的最终数据
 */
struct FoundRecord {
    std::string data_time;   // 消费时间
    std::string car_number;  // 车牌号码
    std::string quantity;    // 数量
    std::string price;       // 单价
    std::string total_cost;  // 总价
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
    std::string count_cost;                 // 该表对应的总金额 
    std::vector<ValueResult> value_results; // 每个源值对应的所有匹配记录
    std::vector<std::string> find_values;   // 该表中所有需要查找的值
};

/**
 * @brief Excel 查找业务处理类
 */
class ExcelFinder
{
private:
    struct ColumnIndex {
        int data_col = -1;
        int car_col  = -1;
        int num_col  = -1;
    };

    struct QuantityRange
    {
        double min_quantity = 0.0;
        double max_quantity = 0.0;

        bool operator<(const QuantityRange& other) const
        {
            if (min_quantity != other.min_quantity)
                return min_quantity < other.min_quantity;
            return max_quantity < other.max_quantity;
        }
    };

    using price = double;
    using PriceMap = std::map<QuantityRange, price>;

public:
    ExcelFinder();

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
    bool export_results();

    // 在控制台打印查找结果
    void print_results() const;

    // 重置结果表（即删除表后重新创建表）
    
    void set_source_path(const std::string& source_path) {
        std::cout << "[ExcelFinder] 设置源文件路径: " << source_path << std::endl;
        source_file_path = source_path;
    }

    void set_target_path(const std::string& target_path) {
        std::cout << "[ExcelFinder] 设置目标文件路径: " << target_path << std::endl;
        target_file_path = target_path;
    }

    void set_output_path(const std::string& output_path) {
        std::cout << "[ExcelFinder] 设置输出文件路径: " << output_path << std::endl;
        output_file_path = output_path;
    }

    bool add_price(double min_quantity, double max_quantity, double price_value);
    bool remove_price(double min, double max);
    bool query_price(double quantity, price& out_price) const;

private:
    XlsxHandle source_xlsx_;
    XlsxHandle target_xlsx_;
    std::string source_file_path = "";
    std::string target_file_path = "";
    std::string output_file_path = "";

    std::string data_tag_;  // 日期标签
    std::string car_tag_;   // 车牌标签
    std::string num_tag_;   // 数量标签
    std::string price_tag_; // 价格标签

    unsigned int read_start_row_ = 2;
    unsigned int read_start_col_ = 1;   // A
    unsigned int read_end_row_ = 9;
    unsigned int read_end_col_ = 7;     // G

    std::vector<SheetResult> results_;

    PriceMap price_map_;

private:
    bool find_and_extract_data_from_target_fast(const std::string& target_sheet_name,
                                                const ColumnIndex& col_idx,
                                                ValueResult& out_value_result,
                                                double &sheet_total_cost);

    void date_simplify(std::string& date_str);

    // 缓存目标表文件
    bool cache_target_sheet(std::string& target_sheet, ColumnIndex& col_idx);
};

