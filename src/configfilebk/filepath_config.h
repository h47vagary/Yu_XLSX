#pragma once

#include <string>

// 必须在 jsoncpp 头之前
#define JSONCPP_DISABLE_STRING_VIEW
#include <json/json.h>

struct FilePathConfig
{
    std::string config_file_path;              // 配置文件路径
    std::string data_source_file_default_path; // 数据源文件默认路径
    std::string search_term_file_default_path; // 搜索项文件默认路径
    std::string export_file_default_path;      // 默认导出路径

    static bool from_file(const std::string& path,
                          FilePathConfig& file_path_config);

    void to_file() const;
};
