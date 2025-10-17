/**
 * @file xlsx_handle.cpp
 * @brief XLSX处理类（基于 OpenXLSX 0.3.x）实现
 * @version 1.0
 * @date 2025-10-10
 * 
 * @copyright Copyright (c) 2025
 */

#include "xlsx_handle.h"
#include <algorithm>
#include <cctype>

XlsxHandle::XlsxHandle()
    : is_open_(false)
    , is_new_file_(false)
    , current_file_path_("")
{
}

XlsxHandle::~XlsxHandle()
{
    close_file();
}

// 打开已有Excel文件
bool XlsxHandle::open_file(const std::string& path)
{
    // 先关闭已打开的文件
    if (is_open_)
    {
        close_file();
    }

    try
    {
        doc_.open(path);
        is_open_ = true;
        is_new_file_ = false;
        current_file_path_ = path;
        std::cout << "[XlsxHandle] Success: Open file - " << path << std::endl;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Open file failed - " << e.what() << std::endl;
        is_open_ = false;
        current_file_path_.clear();
        return false;
    }
}

// 创建新Excel文件
bool XlsxHandle::create_file(const std::string& path)
{
    // 先关闭已打开的文件
    if (is_open_)
    {
        close_file();
    }

    try
    {
        doc_.create(path);
        is_open_ = true;
        is_new_file_ = true;
        current_file_path_ = path;
        std::cout << "[XlsxHandle] Success: Create file - " << path << std::endl;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Create file failed - " << e.what() << std::endl;
        is_open_ = false;
        current_file_path_.clear();
        return false;
    }
}

// 保存文件（空路径 = 保存到当前路径；非空 = 另存为）
bool XlsxHandle::save_file(const std::string& path)
{
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Save failed - no file open!" << std::endl;
        return false;
    }

    try
    {
        if (!path.empty())
        {
            doc_.saveAs(path);
            current_file_path_ = path;
            is_new_file_ = false;
            std::cout << "[XlsxHandle] Success: Save as - " << path << std::endl;
        }
        else
        {
            if (is_new_file_)
            {
                std::cerr << "[XlsxHandle] Error: Save failed - new file needs a path (use save_file(\"path.xlsx\"))!" << std::endl;
                return false;
            }
            doc_.save();
            std::cout << "[XlsxHandle] Success: Save to - " << current_file_path_ << std::endl;
        }
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Save failed - " << e.what() << std::endl;
        return false;
    }
}

// 关闭文件
void XlsxHandle::close_file()
{
    if (is_open_)
    {
        try
        {
            doc_.close();
            std::cout << "[XlsxHandle] Success: Close file - " << current_file_path_ << std::endl;
        }
        catch (const OpenXLSX::XLException& e)
        {
            std::cerr << "[XlsxHandle] Error: Close file failed - " << e.what() << std::endl;
        }
        is_open_ = false;
        is_new_file_ = false;
        current_file_path_.clear();
    }
}

// 判断文件是否已打开
bool XlsxHandle::is_file_open() const
{
    return is_open_;
}

// 获取所有工作表名称
std::vector<std::string> XlsxHandle::get_sheet_names() const
{
    std::vector<std::string> sheet_names;
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Get sheet names failed - no file open!" << std::endl;
        return sheet_names;
    }

    try
    {
        sheet_names = doc_.workbook().worksheetNames();
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Get sheet names failed - " << e.what() << std::endl;
    }
    return sheet_names;
}

// 判断工作表是否存在
bool XlsxHandle::is_sheet_exist(const std::string& sheet_name)
{
    OpenXLSX::XLWorksheet ws;
    return get_worksheet(sheet_name, ws);
}

bool XlsxHandle::create_sheet(const std::string &sheet_name)
{
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Create sheet failed - no file open!" << std::endl;
        return false;
    }
    try
    {
        // 若存在则删除
        delete_sheet(sheet_name);

        doc_.workbook().addWorksheet(sheet_name);
        std::cout << "[XlsxHandle] Success: Create sheet - " << sheet_name << std::endl;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Create sheet failed - " << e.what() << std::endl;
        return false;
    }
}

bool XlsxHandle::delete_sheet(const std::string &sheet_name)
{
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Delete sheet failed - no file open!" << std::endl;
        return false;
    }
    try
    {
        // 判断是否存在
        auto names = doc_.workbook().worksheetNames();
        if (std::find(names.begin(), names.end(), sheet_name) == names.end())
        {
            std::cout << "[XlsxHandle] Info: Delete sheet skipped - '" << sheet_name << "' not exist." << std::endl;
            return true; // 不存在视为成功
        }

        doc_.workbook().deleteSheet(sheet_name);
        std::cout << "[XlsxHandle] Success: Delete sheet - " << sheet_name << std::endl;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Delete sheet failed - " << e.what() << std::endl;
        return false;
    }
}

// 读取指定单元格内容（返回字符串形式）
std::string XlsxHandle::read_cell(const std::string& sheet_name, unsigned int row, unsigned int col)
{
    // 1. 检查文件状态
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Read cell failed - no file open!" << std::endl;
        return "";
    }

    // 2. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return "";
    }

    // 3. 读取单元格
    try
    {
        return ws.cell(row, col).value().getString();
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Read cell (" << row << "," << col << ") failed - " << e.what() << std::endl;
        return "";
    }
}

// 写入字符串到指定单元格
bool XlsxHandle::write_cell(const std::string& sheet_name, unsigned int row, unsigned int col, const std::string& content)
{
    return write_cell_value<std::string>(sheet_name, row, col, content);
}

// 查找指定值在工作表中的所有位置
std::vector<CellPos> XlsxHandle::find_cell_by_value(const std::string& sheet_name, const std::string& target, bool case_sensitive)
{
    std::vector<CellPos> match_positions;
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Find failed - no file open!" << std::endl;
        return match_positions;
    }

    // 1. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return match_positions;
    }

    // 2. 获取已用区域（避免遍历无数据单元格）
    unsigned int start_row, start_col, end_row, end_col;
    if (!get_used_range(sheet_name, start_row, start_col, end_row, end_col))
    {
        std::cerr << "[XlsxHandle] Warning: Find failed - no used range in sheet " << sheet_name << std::endl;
        return match_positions;
    }

    // 3. 定义忽略前后空格接口
    auto trim = [](std::string& s)
    {
        s.erase(s.begin(), std::find_if(s.begin(), s.end(), [](unsigned char ch) {
            return !std::isspace(ch);
        }));
        s.erase(std::find_if(s.rbegin(), s.rend(), [](unsigned char ch) {
            return !std::isspace(ch);
        }).base(), s.end());
    };

    // 4. 处理大小写（不区分则统一转小写）与去空格
    std::string target_str = target;
    trim(target_str);
    if (!case_sensitive)
    {
        std::transform(target_str.begin(), target_str.end(), target_str.begin(), [](unsigned char c) {
            return std::tolower(c);
        });
    }

    // 5. 遍历已用区域查找匹配
    try
    {
        for (unsigned int r = start_row; r <= end_row; ++r)
        {
            for (unsigned int c = start_col; c <= end_col; ++c)
            {
                // 读取单元格内容并处理大小写
                std::string cell_str = ws.cell(r, c).value().getString();
                trim(cell_str);
                if (!case_sensitive)
                {
                    std::transform(cell_str.begin(), cell_str.end(), cell_str.begin(), [](unsigned char c) {
                        return std::tolower(c);
                    });
                }

                // 匹配成功则记录位置
                if (cell_str == target_str)
                {
                    match_positions.emplace_back(r, c);
                }
            }
        }

        // 输出查找结果
        if (match_positions.empty())
        {
            std::cout << "[XlsxHandle] Info: Find '" << target << "' - no match in sheet " << sheet_name << std::endl;
        }
        else
        {
            std::cout << "[XlsxHandle] Success: Find '" << target << "' - " << match_positions.size() << " matches in sheet " << sheet_name << std::endl;
        }
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Find failed - " << e.what() << std::endl;
    }

    return match_positions;
}

// 读取指定行的所有单元格内容
std::vector<std::string> XlsxHandle::read_row(const std::string& sheet_name, unsigned int row)
{
    std::vector<std::string> row_data;
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Read row failed - no file open!" << std::endl;
        return row_data;
    }

    // 1. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return row_data;
    }

    // 2. 获取已用区域（确定列范围）
    unsigned int start_row, start_col, end_row, end_col;
    if (!get_used_range(sheet_name, start_row, start_col, end_row, end_col))
    {
        std::cerr << "[XlsxHandle] Warning: Read row " << row << " failed - no used range in sheet " << sheet_name << std::endl;
        return row_data;
    }

    // 3. 检查目标行是否在已用区域内
    if (row < start_row || row > end_row)
    {
        std::cerr << "[XlsxHandle] Warning: Read row " << row << " failed - row out of used range (" << start_row << "-" << end_row << ")" << std::endl;
        return row_data;
    }

    // 4. 读取该行所有列
    try
    {
        for (unsigned int c = start_col; c <= end_col; ++c)
        {
            row_data.push_back(ws.cell(row, c).value().getString());
        }
        std::cout << "[XlsxHandle] Success: Read row " << row << " - " << row_data.size() << " cells" << std::endl;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Read row " << row << " failed - " << e.what() << std::endl;
    }

    return row_data;
}

// 读取指定列的所有单元格内容
std::vector<std::string> XlsxHandle::read_col(const std::string& sheet_name, unsigned int col)
{
    std::vector<std::string> col_data;
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Read col failed - no file open!" << std::endl;
        return col_data;
    }

    // 1. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return col_data;
    }

    // 2. 获取已用区域（确定行范围）
    unsigned int start_row, start_col, end_row, end_col;
    if (!get_used_range(sheet_name, start_row, start_col, end_row, end_col))
    {
        std::cerr << "[XlsxHandle] Warning: Read col " << col << " failed - no used range in sheet " << sheet_name << std::endl;
        return col_data;
    }

    // 3. 检查目标列是否在已用区域内
    if (col < start_col || col > end_col)
    {
        std::cerr << "[XlsxHandle] Warning: Read col " << col << " failed - col out of used range (" << start_col << "-" << end_col << ")" << std::endl;
        return col_data;
    }

    // 4. 读取该列所有行
    try
    {
        for (unsigned int r = start_row; r <= end_row; ++r)
        {
            col_data.push_back(ws.cell(r, col).value().getString());
        }
        std::cout << "[XlsxHandle] Success: Read col " << col << " - " << col_data.size() << " cells" << std::endl;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Read col " << col << " failed - " << e.what() << std::endl;
    }

    return col_data;
}

std::vector<std::string> XlsxHandle::read_range(const std::string &sheet_name, unsigned int start_row, unsigned int start_col, unsigned int end_row, unsigned int end_col, bool include_empty_cells)
{
    std::vector<std::string> range_data;
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Read range failed - no file open!" << std::endl;
        return range_data;
    }

    // 1. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return range_data;
    }

    // 2. 获取已用区域（确定行列范围）
    unsigned int used_start_row, used_start_col, used_end_row, used_end_col;
    if (!get_used_range(sheet_name, used_start_row, used_start_col, used_end_row, used_end_col))
    {
        // 如果工作表为空，但用户请求的范围合法且希望包含空单元格，则返回相应数量的空字符串
        if (include_empty_cells && start_row <= end_row && start_col <= end_col) {
            size_t num_cells = (end_row - start_row + 1) * (end_col - start_col + 1);
            return std::vector<std::string>(num_cells, "");
        }
        std::cerr << "[XlsxHandle] Warning: Read range failed - no used range in sheet " << sheet_name << std::endl;
        return range_data;
    }

    // 3. 检查目标范围是否在已用区域内
    if (start_row < used_start_row || end_row > used_end_row || start_col < used_start_col || end_col > used_end_col || start_row > end_row || start_col > end_col)
    {
        std::cerr << "[XlsxHandle] Warning: Read range failed - range out of used range (" << used_start_row << "," << used_start_col << ") to (" << used_end_row << "," << used_end_col << ")" << std::endl;
        return range_data;
    }
    
    // 4. 读取该范围所有单元格（按行优先顺序）
    try
    {
        for (unsigned int r = start_row; r <= end_row; ++r)
        {
            for (unsigned int c = start_col; c <= end_col; ++c)
            {
                std::string cell_value = ws.cell(r, c).value().getString();

                // 5. 根据参数决定是否添加空单元格
                if (include_empty_cells || !cell_value.empty())
                {
                    range_data.push_back(cell_value);
                }
            }
        }
        //std::cout << "[XlsxHandle] Success: Read range (" << start_row << "," << start_col << ") to (" << end_row << "," << end_col << ") - " << range_data.size() << " cells" << std::endl;
    }
    catch (const OpenXLSX::XLException& e)
    {
        //std::cerr << "[XlsxHandle] Error: Read range failed - " << e.what() << std::endl;
    }

    return range_data;
}

// 获取工作表的已用区域（包含数据的最小矩形）
bool XlsxHandle::get_used_range(const std::string& sheet_name, unsigned int& start_row, unsigned int& start_col, unsigned int& end_row, unsigned int& end_col)
{
    // 初始化输出参数
    start_row = std::numeric_limits<unsigned int>::max();
    start_col = std::numeric_limits<unsigned int>::max();
    end_row = 0;
    end_col = 0;

    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Get used range failed - no file open!" << std::endl;
        return false;
    }

    // 1. 获取工作表
    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws))
    {
        return false;
    }

    // 2. 遍历合理范围（可根据需求调整 MAX_ROW/MAX_COL，默认 1000 行 + 26 列）
    const unsigned int MAX_ROW = 1000;
    const unsigned int MAX_COL = 26;

    try
    {
        for (unsigned int r = 1; r <= MAX_ROW; ++r)
        {
            for (unsigned int c = 1; c <= MAX_COL; ++c)
            {
                // 跳过空单元格
                if (is_cell_empty(ws, r, c))
                {
                    continue;
                }

                // 更新已用区域范围
                if (r < start_row) start_row = r;
                if (c < start_col) start_col = c;
                if (r > end_row) end_row = r;
                if (c > end_col) end_col = c;
            }
        }

        // 检查是否有数据（无数据则返回 false）
        if (end_row == 0 || end_col == 0)
        {
            std::cerr << "[XlsxHandle] Warning: Get used range failed - sheet " << sheet_name << " is empty" << std::endl;
            return false;
        }

        std::cout << "[XlsxHandle] Success: Get used range - (" << start_row << "," << start_col << ") to (" << end_row << "," << end_col << ")" << std::endl;
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Get used range failed - " << e.what() << std::endl;
        return false;
    }
}

// 获取工作表的最后一行（基于已用区域）
unsigned int XlsxHandle::get_last_row(const std::string& sheet_name)
{
    unsigned int start_row, start_col, end_row, end_col;
    if (get_used_range(sheet_name, start_row, start_col, end_row, end_col))
    {
        return end_row;
    }
    return 0; // 无数据时返回 0
}

// 获取工作表的最后一列（基于已用区域）
unsigned int XlsxHandle::get_last_col(const std::string& sheet_name)
{
    unsigned int start_row, start_col, end_row, end_col;
    if (get_used_range(sheet_name, start_row, start_col, end_row, end_col))
    {
        return end_col;
    }
    return 0; // 无数据时返回 0
}

// 获取整个工作表（基于已用区域）
bool XlsxHandle::cache_sheet_data(const std::string &sheet_name)
{
    if (!is_open_) return false;
    if (sheet_cache_.count(sheet_name)) return true; // 已缓存

    OpenXLSX::XLWorksheet ws;
    if (!get_worksheet(sheet_name, ws)) return false;

    unsigned int start_r, start_c, end_r, end_c;
    if (!get_used_range(sheet_name, start_r, start_c, end_r, end_c)) return false;

    auto trim = [](const std::string &s) -> std::string {
        size_t start = s.find_first_not_of(" \t\r\n");
        if (start == std::string::npos)
            return "";
        size_t end = s.find_last_not_of(" \t\r\n");
        return s.substr(start, end - start + 1);
    };

    std::vector<std::vector<std::string>> table;
    table.reserve(end_r - start_r + 1);

    for (unsigned int r = start_r; r <= end_r; ++r)
    {
        std::vector<std::string> row;
        row.reserve(end_c - start_c + 1);
        for (unsigned int c = start_c; c <= end_c; ++c)
        {
            std::string value = ws.cell(r, c).value().getString();
            row.push_back(trim(value)); // 去除首尾空格
        }
        table.push_back(std::move(row));
    }

    sheet_cache_[sheet_name] = std::move(table);
    range_cache_[sheet_name] = {start_r, start_c, end_r, end_c};
    return true;
}


bool XlsxHandle::get_used_range_cached(const std::string &sheet_name, unsigned int &sr, unsigned int &sc, unsigned int &er, unsigned int &ec)
{
    if (range_cache_.count(sheet_name)) {
        std::tie(sr, sc, er, ec) = range_cache_[sheet_name];
        return true;
    }
    return get_used_range(sheet_name, sr, sc, er, ec);
}

bool XlsxHandle::build_index(const std::string &sheet_name, bool case_sensitive)
{
    if (!cache_sheet_data(sheet_name)) {
        std::cout << "not cache sheet data in build index" << std::endl;
        return false;
    } 

    const auto& table = sheet_cache_[sheet_name];
    std::unordered_map<std::string, std::vector<CellPos>>& index = index_cache_[sheet_name];
    index.clear();

    for (unsigned int r = 0; r < table.size(); ++r) {
        for (unsigned int c = 0; c < table[r].size(); ++c) {
            std::string key = table[r][c];  // 具体单元格的值
            if (!case_sensitive) {
                std::transform(key.begin(), key.end(), key.begin(), ::tolower); // 默认转小写
            }
            index[key].emplace_back(r + 1, c + 1);  // 将行列添加到索引缓存
        }
    }

    std::cout << "[XlsxHandle] Index built for '" << sheet_name 
              << "' (" << index.size() << " unique values)" << std::endl;
    return true;
}

// 第一次搜索：建立索引, 后续搜索：读哈希表返回结果
std::vector<CellPos> XlsxHandle::find_cell_by_value_fast(const std::string &sheet_name, const std::string &value, bool case_sensitive)
{
    if (!index_cache_.count(sheet_name)) {
        build_index(sheet_name, case_sensitive);
    }

    std::string key = value;
    if (!case_sensitive)
        std::transform(key.begin(), key.end(), key.begin(), ::tolower);

    auto it = index_cache_[sheet_name].find(key);
    if (it != index_cache_[sheet_name].end()) {
        return it->second;
    }
    return {};
}

// 辅助函数：获取工作表
bool XlsxHandle::get_worksheet(const std::string& sheet_name, OpenXLSX::XLWorksheet& out_ws)
{
    if (!is_open_)
    {
        std::cerr << "[XlsxHandle] Error: Get worksheet failed - no file open!" << std::endl;
        return false;
    }

    try
    {
        out_ws = doc_.workbook().worksheet(sheet_name);
        return true;
    }
    catch (const OpenXLSX::XLException& e)
    {
        std::cerr << "[XlsxHandle] Error: Get worksheet failed - sheet '" << sheet_name << "' not exist (" << e.what() << ")" << std::endl;
        return false;
    }
}

// 辅助函数：判断单元格是否为空
bool XlsxHandle::is_cell_empty(const OpenXLSX::XLWorksheet& ws, unsigned int row, unsigned int col)
{
    try
    {
        // 字符串为空 = 空单元格
        std::string cell_str = ws.cell(row, col).value().getString();
        return cell_str.empty();
    }
    catch (...)
    {
        // 捕获所有异常（如单元格是数字类型），视为非空
        return false;
    }
}