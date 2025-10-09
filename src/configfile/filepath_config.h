/**
 * @file filepath_config.h
 * @brief 默认文件路径配置文件
 * @version 0.1
 * @date 2025-10-09
 * 
 * @copyright Copyright (c) 2025
 * 
 */

#pragma once
#include <string>
#include "json/json.h"

struct FilePathConfig
{
    std::string data_source_file_default_path;      // 数据源文件默认路径
    std::string search_term_file_default_path;      // 搜索项文件默认路径
    std::string export_file_default_path;           // 默认导出路径

    static bool from_file(const std::string& path, FilePathConfig& file_path_config);
    void to_file(const std::string& path) const;
};