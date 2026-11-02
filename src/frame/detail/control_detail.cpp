// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/detail/control_detail.h"

#include "frame/window_impl.h"

#include <algorithm>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <functional>
#include <limits>
#include <commdlg.h>
#include <pathcch.h>
#include <richedit.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <windowsx.h>
#include <winternl.h>
#include "ui/feedback.h"
#include "frame/commands/command_ids.h"
#include "registry/registry_store.h"
#include "dialogs/security_dialog.h"
#include "defaults/default_loader.h"
#include "dialogs/comment_editor.h"
#include "dialogs/value_editor.h"
#include "frame/window/message_dispatch.h"
#include "frame/window/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"

#include <win32ui/system/date_time.h>

namespace regkit::window_detail
{

std::wstring LeafName(const RegistryNode& node)
{
    if (node.subkey.empty())
    {
        return node.root_name.empty() ? registry_path::RootName(node.root) : node.root_name;
    }
    return registry_path::Leaf(node.subkey);
}

bool UseBinaryValueIcon(DWORD type)
{
    switch (type)
    {
    case REG_NONE:
    case REG_BINARY:
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
    case REG_QWORD:
    case REG_RESOURCE_LIST:
    case REG_FULL_RESOURCE_DESCRIPTOR:
    case REG_RESOURCE_REQUIREMENTS_LIST:
    case REG_LINK:
        return true;
    default:
        return false;
    }
}

ListRow MakeValueListRow(const std::wstring& name, DWORD type, const BYTE* data, DWORD data_size, bool preview)
{
    ListRow row;
    row.name = name.empty() ? util::Tr(L"(Default)") : registry_path::DisplayName(name);
    row.type = value_format::TypeName(type);
    row.data_ready = data_size == 0 || data != nullptr;
    if (row.data_ready && data_size > 0)
    {
        // preview keeps displayed text, filters reload full data
        row.data_preview = preview && data_size > kValuePreviewBytes;
        row.data = value_format::DisplayData(type, data, row.data_preview ? kValuePreviewBytes : data_size);
        if (row.data_preview && row.data.size() > kValuePreviewLimit)
        {
            row.data.resize(kValuePreviewLimit);
        }
    }
    row.image_index = UseBinaryValueIcon(type) ? kBinaryIconIndex : kValueIconIndex;
    row.kind = rowkind::kValue;
    row.extra = name;
    row.size = std::to_wstring(data_size);
    row.size_value = data_size;
    row.has_size = true;
    row.value_type = type;
    row.value_data_size = data_size;
    return row;
}

void UpdateLeafName(RegistryNode* node, const std::wstring& new_name)
{
    if (!node || node->subkey.empty())
    {
        return;
    }
    size_t pos = node->subkey.rfind(L'\\');
    if (pos == std::wstring::npos)
    {
        node->subkey = new_name;
    }
    else
    {
        node->subkey = node->subkey.substr(0, pos + 1) + new_name;
    }
}

std::wstring FormatFileTime(const FILETIME& filetime)
{
    if (filetime.dwLowDateTime == 0 && filetime.dwHighDateTime == 0)
    {
        return L"";
    }
    return win32ui::FormatFileTime(filetime);
}

std::wstring FormatCommentDisplay(const std::wstring& text)
{
    std::wstring out = win32ui::SingleLineText(text);
    out.erase(std::unique(out.begin(), out.end(), [](wchar_t left, wchar_t right) { return left == L' ' && right == L' '; }), out.end());
    return out;
}

static int CompareUint64(uint64_t left, uint64_t right)
{
    return (left > right) - (left < right);
}

const std::wstring& ValueRowFieldText(const ListRow& row, int subitem)
{
    switch (subitem)
    {
    case kValueColName:
        return row.name;
    case kValueColType:
        return row.type;
    case kValueColData:
        return row.data;
    case kValueColDefault:
        return row.default_data;
    case kValueColReadOnBoot:
        return row.read_on_boot;
    case kValueColSize:
        return row.size;
    case kValueColDate:
        return row.date;
    case kValueColDetails:
        return row.details;
    case kValueColComment:
        return row.comment;
    default:
        return row.extra;
    }
}

static int CompareValueRow(const ListRow& left, const ListRow& right, int column)
{
    switch (column)
    {
    case kValueColName:
        return util::CompareListText(left.name, right.name);
    case kValueColType:
        return util::CompareListText(left.type, right.type);
    case kValueColData:
        return util::CompareListText(left.data, right.data);
    case kValueColDefault:
        return util::CompareListText(left.default_data, right.default_data);
    case kValueColReadOnBoot:
        return util::CompareListText(left.read_on_boot, right.read_on_boot);
    case kValueColSize:
        if (left.has_size != right.has_size)
        {
            return left.has_size ? -1 : 1;
        }
        return CompareUint64(left.size_value, right.size_value);
    case kValueColDate:
        if (left.has_date != right.has_date)
        {
            return left.has_date ? -1 : 1;
        }
        return CompareUint64(left.date_value, right.date_value);
    case kValueColDetails:
        if (left.has_details != right.has_details)
        {
            return left.has_details ? -1 : 1;
        }
        if (left.detail_key_count != right.detail_key_count)
        {
            return CompareUint64(left.detail_key_count, right.detail_key_count);
        }
        return CompareUint64(left.detail_value_count, right.detail_value_count);
    case kValueColComment:
        return util::CompareListText(left.comment, right.comment);
    default:
        return util::CompareListText(left.name, right.name);
    }
}

void SortValueRows(std::vector<ListRow>* rows, int column, bool ascending, const util::SortKeys& earlier)
{
    if (!rows || rows->size() < 2)
    {
        return;
    }
    std::stable_sort(rows->begin(), rows->end(), [&](const ListRow& left, const ListRow& right) {
        if (left.kind != right.kind)
        {
            return left.kind == rowkind::kKey;
        }
        return util::CompareSorted(column, ascending, earlier, [&](int key) { return CompareValueRow(left, right, key); }) < 0;
    });
}

HFONT CreateUIFont()
{
    return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

void ApplyFont(HWND hwnd, HFONT font)
{
    if (hwnd && font)
    {
        SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
}

HTREEITEM FindChildByText(HWND tree, HTREEITEM parent, const std::wstring& text)
{
    return win32ui::FindTreeChild(tree, parent, registry_path::DisplayName(text));
}

}
