#include "config_manager.h"

ConfigManager& ConfigManager::instance()
{
    static ConfigManager inst;
    return inst;
}

void ConfigManager::register_configurable(IConfigurable* obj)
{
    std::lock_guard<std::mutex> lock(mutex_);
    objects_[obj->config_key()] = obj;
}

bool ConfigManager::load_all()
{
    nlohmann::json root = load_file();

    for (auto& [key, obj] : objects_)
    {
        if (root.contains(key))
            obj->load_config(root[key]);
    }
    return true;
}

bool ConfigManager::save_all()
{
    nlohmann::json root = nlohmann::json::object();

    for (auto& [key, obj] : objects_)
    {
        root[key] = obj->dump_config();
    }
    save_file(root);
    return true;
}

void ConfigManager::set_config_file(const std::string &path)
{
    std::lock_guard<std::mutex> lock(mutex_);
    config_file_ = path;
}

nlohmann::json ConfigManager::load_file()
{
    std::lock_guard<std::mutex> lock(mutex_);

    std::ifstream ifs(config_file_);
    if (!ifs.is_open())
    {
        // 文件不存在，返回空对象
        return nlohmann::json::object();
    }

    nlohmann::json root;
    try {
        ifs >> root;
    }
    catch (const std::exception& e)
    {
        return nlohmann::json::object();
    }

    return root;
}

void ConfigManager::save_file(const nlohmann::json &root)
{
    std::lock_guard<std::mutex> lock(mutex_);

    std::string tmp_file = config_file_ + ".tmp";
    std::ofstream ofs(tmp_file, std::ios::trunc);
    if (!ofs.is_open())
        return;

    ofs << root.dump(4) << std::endl;
    ofs.close();

    // 原子替换
    std::remove(config_file_.c_str());
    std::rename(tmp_file.c_str(), config_file_.c_str());
}
