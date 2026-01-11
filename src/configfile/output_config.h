#pragma once

#include <string>
#include <map>

// 必须在 jsoncpp 头之前
#define JSONCPP_DISABLE_STRING_VIEW
#include <json/json.h>

struct OutputConfig
{
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

    std::string config_file_path;
    PriceMap price_map;

    OutputConfig();

    // === 业务接口 ===
    bool add_price(double min, double max, price value);
    bool remove_price(double min, double max);
    bool query_price(double quantity, price& out_price) const;
    void clear();

    // === 持久化 ===
    static bool from_file(const std::string& path,
                          OutputConfig& output_config);
    void to_file() const;
};
