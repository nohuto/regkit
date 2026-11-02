// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "dialogs/bitfield_definition.h"

namespace regkit::editors
{

bool EditBitfieldDefinition(HWND owner, bitfield::Definition* definition, bool lock_width);
bool EditBitfieldField(HWND owner, const bitfield::Definition& parent, int editing, bitfield::Field* field);
void ShowBitfieldDefinitionEditor(HWND owner, const std::wstring& path);

}
