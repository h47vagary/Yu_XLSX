#include "time_utils.h"

#include <unistd.h>
#include <fstream>
#include <string>

uint64_t TimeDealUtils::get_timestamp()
{
    // 获取当前时间点
    auto now = std::chrono::system_clock::now();

    // 获取秒级时间戳
    std::time_t now_time_t = std::chrono::system_clock::to_time_t(now);

    // 计算毫秒部分
    auto duration = now.time_since_epoch();
    uint64_t milliseconds = std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
    uint64_t now_seconds = std::chrono::duration_cast<std::chrono::seconds>(duration).count();
    uint64_t millis_part = milliseconds - now_seconds * 1000;

    // 返回完整的 Unix 时间戳（单位：毫秒）
    return static_cast<uint64_t>(now_time_t) * 1000 + millis_part;
}

uint64_t TimeDealUtils::get_time_difference(uint64_t start, uint64_t end)
{
    return (end > start) ? (end - start) : 0;
}

std::string TimeDealUtils::timestamp_to_string(uint64_t timestamp)
{
    std::time_t time = static_cast<std::time_t>(timestamp / 1000);
    std::tm tm_time;

#ifdef _WIN32
    localtime_s(&tm_time, &time);
#else
    localtime_r(&time, &tm_time);
#endif

    char buffer[6];
    std::sprintf(buffer, "%02d:%02d", tm_time.tm_min, tm_time.tm_sec);

    return std::string(buffer);
}