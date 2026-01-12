#pragma once

#include <string>
#include "nlohmann/json.hpp"

/**
 * @brief 配置接口
 * @details 业务类只继承实现此类，只接触 Json::Value，不关心文件，不关心路径
 */
class IConfigurable
{
public:
    virtual ~IConfigurable() = default;

    // 配置唯一 key
    virtual std::string config_key() const = 0;

    // 导出配置
    virtual nlohmann::json dump_config() const = 0;

    // 加载配置
    virtual void load_config(const Json::Value& cfg) = 0;
};