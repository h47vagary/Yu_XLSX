#include "filepath_config.h"
#include "config_file.h"

#include <iostream>

bool FilePathConfig::from_file(const std::string& path,
                               FilePathConfig& file_path_config)
{
    ConfigFile file;
    if (!file.load(path))
    {
        std::cerr << "Failed to load fileconfig config file: "
                  << path << std::endl;
        return false;
    }

    const Json::Value& root = file.root();

    READ_IF_MEMBER(root,
                   "data_source_file_default_path",
                   file_path_config.data_source_file_default_path,
                   String);

    READ_IF_MEMBER(root,
                   "search_term_file_default_path",
                   file_path_config.search_term_file_default_path,
                   String);

    READ_IF_MEMBER(root,
                   "export_file_default_path",
                   file_path_config.export_file_default_path,
                   String);

    return true;
}

void FilePathConfig::to_file() const
{
    ConfigFile file;

    Json::Value& root = file.root();

    root["data_source_file_default_path"] = data_source_file_default_path;
    root["search_term_file_default_path"] = search_term_file_default_path;
    root["export_file_default_path"]      = export_file_default_path;

    if (!file.save(config_file_path))
    {
        std::cerr << "Failed to save fileconfig config file: "
                  << config_file_path << std::endl;
    }
}
