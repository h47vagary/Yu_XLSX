#pragma once
#include "filepath_config.h"
#include "task_executor.h"
#include <functional>


class FilePathConfigAsync
{
public:
    // 异步读取配置
    static void load_async(const std::string& path,
                            FilePathConfig* config_ptr = nullptr,
                            std::function<void(FilePathConfig)> on_success = nullptr,
                            std::function<void()> on_fail = nullptr)
    {
        TaskExecutor::run([=]() 
        {
            FilePathConfig temp_config;
            if (FilePathConfig::from_file(path, temp_config))
            {
                if (config_ptr) 
                {
                    *config_ptr = std::move(temp_config);
                } 
                else if (on_success) 
                {
                    on_success(std::move(temp_config));
                }
            }
            else
            {
                if (on_fail) on_fail();
            }
        });
    }

    // 异步保存配置
    static void save_async(const std::string& path,
                           FilePathConfig* config_ptr,
                           std::function<void()> on_success = nullptr,
                           std::function<void()> on_fail = nullptr)
    {
        TaskExecutor::run([=]() 
        {
            try {
                config_ptr->to_file();
                if (on_success) on_success();
            } catch (...) {
                if (on_fail) on_fail();
            }
        });
    }
};