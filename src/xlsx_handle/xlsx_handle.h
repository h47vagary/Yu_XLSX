/**
 * @file xlsx_handle.h
 * @brief XLSX处理类（基于 OpenXLSX）
 * @version 0.1
 * @date 2025-10-09
 * 
 * @copyright Copyright (c) 2025
 * 
 */

#include "OpenXLSX.hpp"
#include <string>
#include <vector>


class XlsxHandle
{
public:
    XlsxHandle ();
    ~XlsxHandle ();

    /**
     * @brief 打开 Excel 文件
     * @param path 文件路径
     * @return true 成功
     * @return false 失败
     */
    bool open_file (const std::string &path);

    /**
     * @brief 关闭 Excel 文件
     */
    void close_file ();

    /**
     * @brief 获取所有工作表名称
     * @return 工作表名称列表
     */
    std::vector<std::string> get_sheet_names () const;

    /**
     * @brief 读取指定工作表中指定单元格内容
     * @param sheet_name 工作表名
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @return 单元格内容
     */
    std::string read_cell (const std::string &sheet_name, unsigned int row, unsigned int col);

private:
    OpenXLSX::XLDocument doc_;
    bool is_open_;
};