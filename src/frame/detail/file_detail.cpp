// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/detail/file_detail.h"

#include "frame/detail/draw_detail.h"
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
#include "defaults/default_loader.h"
#include "dialogs/comment_editor.h"
#include "dialogs/security_dialog.h"
#include "dialogs/value_editor.h"
#include "frame/commands/command_ids.h"
#include "frame/window/message_dispatch.h"
#include "frame/window/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "ui/feedback.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"

namespace regkit::window_detail
{

bool IsKnownIconSetName(const std::wstring& value)
{
    return util::EqualsInsensitive(value, kIconSetClassic) || util::EqualsInsensitive(value, kIconSetPhosphor) ||
           util::EqualsInsensitive(value, kIconSetCustom);
}

// assets\icons in the folder or one of its parents
static std::wstring FindIconsBelow(std::wstring base, int levels)
{
    for (int i = 0; i < levels && !base.empty(); ++i)
    {
        std::wstring candidate = util::JoinPath(base, L"assets\\icons");
        if (util::IsDirectory(candidate.c_str()))
        {
            return candidate;
        }
        std::wstring parent = win32ui::ParentDirectory(base);
        if (parent == base)
        {
            break;
        }
        base = std::move(parent);
    }
    return L"";
}

static std::wstring FindAssetsIconsRoot()
{
    std::wstring found = FindIconsBelow(util::ModuleDirectory(), 6);
    return found.empty() ? FindIconsBelow(win32ui::FullPath(L"."), 3) : found;
}

std::wstring AssetsIconsRoot()
{
    static std::wstring cached;
    static bool cached_set = false;
    if (!cached_set)
    {
        cached = FindAssetsIconsRoot();
        cached_set = true;
    }
    return cached;
}

std::wstring NormalizeMachineName(const std::wstring& text)
{
    std::wstring trimmed = TrimWhitespace(text);
    while (!trimmed.empty() && (trimmed.back() == L'\\' || trimmed.back() == L'/'))
    {
        trimmed.pop_back();
    }
    if (trimmed.empty())
    {
        return trimmed;
    }
    if (trimmed.rfind(L"\\\\", 0) == 0)
    {
        return trimmed;
    }
    return L"\\\\" + trimmed;
}

std::wstring StripMachinePrefix(const std::wstring& machine)
{
    if (machine.rfind(L"\\\\", 0) == 0)
    {
        return machine.substr(2);
    }
    return machine;
}

bool ReadActiveEntries(const std::wstring& path, std::vector<std::wstring>* entries, std::wstring* error)
{
    std::wstring content;
    if (!util::ReadStateFile(path, &content, error) || content.empty())
    {
        return false;
    }
    for (const std::wstring_view raw : win32ui::LineViews(content))
    {
        std::wstring line = util::TrimWhitespace(raw);
        if (line.empty() || line.front() == L'#')
        {
            continue;
        }
        entries->push_back(std::move(line));
    }
    return true;
}

bool ResolveActiveSource(const std::wstring& entry, const std::function<std::wstring(const std::wstring&)>& resolve_bundled, const wchar_t* fallback_label, std::wstring* source, std::wstring* label)
{
    *source = entry;
    label->clear();
    if (!util::IsFile(source->c_str()))
    {
        const std::wstring bundled = resolve_bundled(entry);
        if (bundled.empty() || !util::IsFile(bundled.c_str()))
        {
            return false;
        }
        *source = bundled;
        *label = entry;
    }
    if (label->empty())
    {
        *label = util::FileBaseName(*source);
    }
    if (label->empty())
    {
        *label = fallback_label;
    }
    return true;
}

}
