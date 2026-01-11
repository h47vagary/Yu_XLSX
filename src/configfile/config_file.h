#pragma once

#include <fstream>
#include <iostream>
#include <string>
#include <sys/stat.h>
#include <sys/types.h>

#ifdef _WIN32
#include <direct.h>
#endif

// ⚠️ 必须在 include jsoncpp 之前
#define JSONCPP_DISABLE_STRING_VIEW
#include <json/json.h>

#define D_CONFIG_BASE_PATH   "./config"
#define D_CONFIG_FILES_PATH "./config/fileconfig.json"

// 更安全的宏：强制 key 为 const char*
#define READ_IF_MEMBER(json, key, var, type)        \
    do {                                            \
        if ((json).isMember(key))                   \
            (var) = (json)[key].as##type();         \
    } while (0)

class ConfigFile
{
public:
    ConfigFile();

    bool load(const std::string& file_path);
    bool save(const std::string& file_path) const;

    void set(const std::string& key, const std::string& value);
    std::string get(const std::string& key) const;
    std::string get_or(const std::string& key,
                       const std::string& default_val) const;

    Json::Value& root();
    const Json::Value& root() const;

private:
    static bool ensure_directory_exists(const std::string& dir);

    Json::Value root_;
};
