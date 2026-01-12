#pragma once

#include "configurable.h"
#include "excel_finder.h"


class ExcelFinderConfigHandler : public IConfigurable
{
public:
    explicit ExcelFinderConfigHandler(ExcelFinder& finder);

    std::string config_key() const override
    {
        return "ExcelFinderConfig";
    }
    nlohmann::json dump_config() const override;
    void load_config(const nlohmann::json &cfg) override;

private:
    ExcelFinder& finder_;
};