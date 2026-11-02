// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <string>

namespace regkit::editors
{

struct CommentScope
{
    bool rule = false;
    bool same_type = true;
    bool same_size = false;
    bool in_key = false;
    bool include_subkeys = false;
    std::wstring key_path;
};

struct CommentRequest
{
    std::wstring text;
    CommentScope scope;
    std::wstring name;
    std::wstring type;
    std::wstring size;
    bool multiple = false;
    bool can_restore = false;
    bool key = false;
};

struct CommentResult
{
    std::wstring text;
    CommentScope scope;
    bool restore_default = false;
};

bool EditComment(HWND owner, const CommentRequest& request, CommentResult* result);

}
