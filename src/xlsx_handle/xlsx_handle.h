/**
 * @file xlsx_handle.h
 * @brief XLSX处理类（基于 OpenXLSX 0.3.x 兼容版）
 * @version 1.0
 * @date 2025-10-10
 * 
 * @copyright Copyright (c) 2025
 * 
 */

#include "OpenXLSX.hpp"
#include <string>
#include <vector>
#include <iostream>
#include <limits>

// 单元格位置结构体（行、列均从 1 开始计数）
struct CellPos
{
    unsigned int row; // 行号
    unsigned int col; // 列号

    CellPos(unsigned int r = 0, unsigned int c = 0) : row(r), col(c) {}

    bool operator==(const CellPos& other) const
    {
        return row == other.row && col == other.col;
    }
};

inline std::ostream& operator<<(std::ostream& os, const CellPos& pos)
{
    os << "Row:" << pos.row << ", Col:" << pos.col;
    return os;
}

class XlsxHandle
{
public:
    XlsxHandle();
    ~XlsxHandle();

    // 核心文件操作
    bool open_file(const std::string& path);        // 打开已有文件
    bool create_file(const std::string& path);      // 创建新文件
    bool save_file(const std::string& path = "");   // 保存文件（支持另存为）
    void close_file();                              // 关闭文件
    bool is_file_open() const;                      // 判断文件是否已打开

    // 工作表操作
    std::vector<std::string> get_sheet_names() const;       // 获取所有工作表名称
    bool is_sheet_exist(const std::string& sheet_name);     // 判断工作表是否存在
    bool create_sheet(const std::string& sheet_name);       // 创建新工作表
    bool delete_sheet(const std::string& sheet_name);       // 删除工作表

    // 单元格读写
    std::string read_cell(const std::string& sheet_name, unsigned int row, unsigned int col);                           // 读单元格
    bool write_cell(const std::string& sheet_name, unsigned int row, unsigned int col, const std::string& content);     // 写字符串
    template<typename T> // 模板写（支持 int/double/bool 等基础类型）
    bool write_cell_value(const std::string& sheet_name, unsigned int row, unsigned int col, const T& value);

    // 单元格查找（忽略前后空格）
    std::vector<CellPos> find_cell_by_value(const std::string& sheet_name, const std::string& target, bool case_sensitive = true); // 查找值

    // 批量读取
    std::vector<std::string> read_row(const std::string& sheet_name, unsigned int row); // 读整行
    std::vector<std::string> read_col(const std::string& sheet_name, unsigned int col); // 读整列
    std::vector<std::string> read_range(const std::string& sheet_name, unsigned int start_row, 
                                        unsigned int start_col, unsigned int end_row, unsigned int end_col,
                                        bool include_empty_cells = false); // 读区域
    

    // 范围获取
    bool get_used_range(const std::string& sheet_name, unsigned int& start_row, unsigned int& start_col, unsigned int& end_row, unsigned int& end_col); // 获取已用区域
    unsigned int get_last_row(const std::string& sheet_name); // 获取最后一行
    unsigned int get_last_col(const std::string& sheet_name); // 获取最后一列
    bool cache_sheet_data(const std::string& sheet_name);     // 获取整个表
    bool get_used_range_cached(const std::string& sheet_name, unsigned int& sr, unsigned int& sc, unsigned int& er, unsigned int& ec);


    // 构建索引
    bool build_index(const std::string& sheet_name, bool case_sensitive);

    // 优化查找函数(O(1)复杂度)
    std::vector<CellPos> find_cell_by_value_fast(const std::string& sheet_name, const std::string& value, bool case_sensitive);

    // 获取缓存表
    const std::vector<std::vector<std::string>>& get_sheet_cache(const std::string& sheet_name) const
    {
        auto it = sheet_cache_.find(sheet_name);
        if (it == sheet_cache_.end())
            throw std::runtime_error("Sheet not cached: " + sheet_name);
        return it->second;
    }
    
    // ================== 样式设置接口 ==================
    void set_font_size(const std::string& sheet_name, int row, int col, double size);
    void set_bold(const std::string& sheet_name, int row, int col, bool bold);


private:
    // 辅助函数
    bool get_worksheet(const std::string& sheet_name, OpenXLSX::XLWorksheet& out_ws); // 获取工作表（输出参数）
    bool is_cell_empty(const OpenXLSX::XLWorksheet& ws, unsigned int row, unsigned int col); // 判断单元格是否为空

    // 成员变量
    OpenXLSX::XLDocument doc_;              // OpenXLSX 文档对象
    bool is_open_;                          // 文件是否已打开
    bool is_new_file_;                      // 是否为新建文件（未保存过）
    std::string current_file_path_;         // 当前文件路径

    // 缓存：sheet_name -> (二维表格数据)
    std::unordered_map<std::string, std::vector<std::vector<std::string>>> sheet_cache_;

    // 缓存：sheet_name -> 哈希索引 (value -> 所有位置)
    std::unordered_map<std::string, std::unordered_map<std::string, std::vector<CellPos>>> index_cache_;

    // 缓存已用范围
    std::unordered_map<std::string, std::tuple<unsigned int, unsigned int, unsigned int, unsigned int>> range_cache_;

};

template<typename T>
bool XlsxHandle::write_cell_value(const std::string& sheet_name, unsigned int row, unsigned int col, const T& value)
{
    // 1. 检查文件是否打开
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Write failed - no file open!" << std::endl;
        return false;
    }

    // 2. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return false;
    }

    // 3. 写入数据（旧版本通过 value() 赋值）
    try
    {
        ws.cell(row, col).value() = value;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Write cell (" << row << "," << col << ") failed - " << e.what() << std::endl;
        return false;
    }
}