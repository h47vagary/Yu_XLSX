#include "output_config.h"
#include <iostream>
#include "config_file.h"

bool OutputConfig::from_file(const std::string &path, OutputConfig &output_config)
{
    ConfigFile file;
    if (!file.load(path))
    {
        std::cerr << "Failed to load output config file: " << path << std::endl;
        return false;
    }

    Json::Value root = file.root();
    if (!root.isMember("prices") || !root["prices"].isArray())
    {
        std::cerr << "Invalid output config format" << std::endl;
        return false;
    }

    output_config.price_map.clear();

    for (const auto& item : root["prices"])
    {
        if (!item.isMember("min") || !item.isMember("max") || !item.isMember("price"))
            continue;

        OutputConfig::QuantityRange range;
        range.min_quantity = item["min"].asDouble();
        range.max_quantity = item["max"].asDouble();

        double price = item["price"].asDouble();

        output_config.price_map.emplace(range, price);
    }

    return true;
}

void OutputConfig::to_file() const
{
    ConfigFile file;
    Json::Value root;
    Json::Value prices(Json::arrayValue);

    for (const auto& [range, price] : price_map)
    {
        Json::Value item;
        item["min"] = range.min_quantity;
        item["max"] = range.max_quantity;
        item["price"] = price;
        prices.append(item);
    }

    root["prices"] = prices;
    file.root() = root;

    if (!file.save(config_file_path))
    {
        std::cerr << "Failed to save output config file: " << config_file_path << std::endl;
    }
}

bool OutputConfig::add_price(double min, double max, price value)
{
    if (min >= max)
        return false;

    QuantityRange new_range{min, max};

    // 区间重叠校验
    for (const auto& [range, _] : price_map)
    {
        bool overlap =
            !(max <= range.min_quantity ||
              min >= range.max_quantity);

        if (overlap)
            return false;
    }

    auto [it, inserted] = price_map.emplace(new_range, value);
    return inserted;
}

bool OutputConfig::remove_price(double min, double max)
{
    QuantityRange range{min, max};
    return price_map.erase(range) > 0;
}

bool OutputConfig::query_price(double quantity, price& out_price) const
{
    for (const auto& [range, price] : price_map)
    {
        if (quantity >= range.min_quantity &&
            quantity <  range.max_quantity)
        {
            out_price = price;
            return true;
        }
    }
    return false;
}

void OutputConfig::clear()
{
    price_map.clear();
}
