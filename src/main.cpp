#include <iostream>
#include <windows.h>
#include "excel_finder.h"

int main() {
    SetConsoleOutputCP(CP_UTF8); // 设置控制台输出为 UTF-8

    // 1. 定义文件路径和业务参数
    const std::string source_file = "../doc/福邦1.xlsx";
    const std::string target_file = "../doc/福邦2-4交易详情135.xlsx";
    const std::string output_file = "../doc/find_result.xlsx";

    // 2. 创建并初始化业务对象
    ExcelFinder finder(source_file, target_file);
    if (!finder.init()) {
        std::cerr << "业务初始化失败，程序退出。" << std::endl;
        return 1;
    }

    // 3. 配置业务参数
    finder.set_tags("消费时间", "车牌号码", "数量");
    finder.set_source_read_range(2, 1, 9, 7);

    // 4. 执行核心业务逻辑
    if (!finder.execute()) {
        std::cerr << "查找任务执行失败。" << std::endl;
        return 1;
    }

    // 5. 处理结果（打印和导出）
    finder.print_results();
    if (!finder.export_results(output_file)) {
        std::cerr << "结果导出失败。" << std::endl;
    }

    std::cout << "\n程序执行完毕。" << std::endl;

    return 0;
}