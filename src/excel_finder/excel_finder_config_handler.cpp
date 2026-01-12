#include "excel_finder_config_handler.h"
#include <vector>

ExcelFinderConfigHandler::ExcelFinderConfigHandler(ExcelFinder& finder)
    : finder_(finder)
{
}

nlohmann::json ExcelFinderConfigHandler::dump_config() const
{
    nlohmann::json cfg;
    cfg["source_file_path"] = finder_.get_source_path();
    cfg["target_file_path"] = finder_.get_target_path();
    cfg["output_file_path"] = finder_.get_output_path();

    std::map<ExcelFinder::QuantityRange, ExcelFinder::price> price_map = finder_.get_price_map();
    std::vector<nlohmann::json> price_list;
    for (const auto& [range, price_value] : price_map)
    {
        nlohmann::json price_entry;
        price_entry["min_quantity"] = range.min_quantity;
        price_entry["max_quantity"] = range.max_quantity;
        price_entry["price"] = price_value;
        price_list.push_back(price_entry);
    }
    cfg["price_list"] = price_list;

    return cfg;
}

void ExcelFinderConfigHandler::load_config(const nlohmann::json &cfg)
{
    if (cfg.contains("source_file_path"))
    {
        std::string source_path = cfg["source_file_path"];
        finder_.set_source_path(source_path);
    }
    if (cfg.contains("target_file_path"))
    {
        std::string target_path = cfg["target_file_path"];
        finder_.set_target_path(target_path);
    }
    if (cfg.contains("output_file_path"))
    {
        std::string output_path = cfg["output_file_path"];
        finder_.set_output_path(output_path);
    }

    if (cfg.contains("price_list") && cfg["price_list"].is_array())
    {
        finder_.clear_price();
        for (const auto& price_entry : cfg["price_list"])
        {
            if (price_entry.contains("min_quantity") &&
                price_entry.contains("max_quantity") &&
                price_entry.contains("price"))
            {
                double min_quantity = price_entry["min_quantity"];
                double max_quantity = price_entry["max_quantity"];
                double price_value = price_entry["price"];
                finder_.add_price(min_quantity, max_quantity, price_value);
            }
        }
    }
}
