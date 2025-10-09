#pragma once

#include <cstdint>
#include <chrono>
#include <iostream>

#if _WIN32
#include <windows.h>
#endif

#define TIMER(Run) [&](){\
    auto t1 = std::chrono::system_clock::now();\
    auto res = Run;\
    auto t2 = std::chrono::system_clock::now();\
    auto dt = std::chrono::duration_cast<std::chrono::milliseconds>(t2 - t1).count();\
    std::cout << #Run << " elapsed time: " << dt << "ms" << std::endl;\
    return res;\
}()

namespace TimeDealUtils
{
    uint64_t get_timestamp();
    uint64_t get_time_difference(uint64_t start, uint64_t end);
    std::string timestamp_to_string(uint64_t timestamp);
}