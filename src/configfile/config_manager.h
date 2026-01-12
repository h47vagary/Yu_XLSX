#pragma once

#include <unordered_map>
#include <string>
#include <mutex>
#include <fstream>
#include "nlohmann/json.hpp"
#include "configurable.h"

class ConfigManager
{
public:
    static ConfigManager& instance();

    void register_configurable(IConfigurable* obj);

    bool load_all();
    bool save_all();

    void set_config_file(const std::string& path);

private:
    ConfigManager() = default;

    nlohmann::json load_file();
    void save_file(const nlohmann::json& root);

private:
    std::unordered_map<std::string, IConfigurable*> objects_;
    std::mutex mutex_;

    std::string config_file_{"config.json"};
};