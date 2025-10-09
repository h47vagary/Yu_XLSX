#include "xlsx_handle.h"
#include <iostream>

using namespace OpenXLSX;

XlsxHandle::XlsxHandle ()
    : is_open_(false)
{
}

XlsxHandle::~XlsxHandle ()
{
    if (is_open_)
        close_file ();
}

bool XlsxHandle::open_file (const std::string &path)
{
    try
    {
        doc_.open (path);
        is_open_ = true;
        return true;
    }
    catch (const std::exception &e)
    {
        std::cerr << "打开文件失败: " << e.what () << std::endl;
        is_open_ = false;
        return false;
    }
}

void XlsxHandle::close_file ()
{
    if (is_open_)
    {
        doc_.close ();
        is_open_ = false;
    }
}

std::vector<std::string> XlsxHandle::get_sheet_names () const
{
    std::vector<std::string> names;
    if (!is_open_)
        return names;

    try
    {
        auto workbook = doc_.workbook ();
        auto sheets = workbook.worksheetNames ();
        for (const auto &name : sheets)
        {
            names.push_back (name);
        }
    }
    catch (const std::exception &e)
    {
        std::cerr << "读取工作表名称失败: " << e.what () << std::endl;
    }

    return names;
}

std::string XlsxHandle::read_cell (const std::string &sheet_name, unsigned int row, unsigned int col)
{
    if (!is_open_)
        return "";

    try
    {
        auto ws = doc_.workbook ().worksheet (sheet_name);
        return ws.cell (XLCellReference (row, col)).value ().getString();
    }
    catch (const std::exception &e)
    {
        std::cerr << "读取单元格失败: " << e.what () << std::endl;
        return "";
    }
}
