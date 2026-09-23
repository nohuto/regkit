// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"
#include "frame/window_impl.h"

#include "editors/bitfield_definition_editor.h"
#include "editors/decoder_dialog.h"
#include "frame/key_handles_window.h"

namespace regkit
{
using namespace command_detail;

bool MainWindow::Impl::HandleToolsCommand(int command_id)
{
    if (command_id >= cmd::kToolsBitfieldFileBase && command_id <= cmd::kToolsBitfieldFileMax)
    {
        const std::vector<editors::bitfield::DefinitionFile>& files = editors::bitfield::BundledFiles();
        const size_t index = static_cast<size_t>(command_id - cmd::kToolsBitfieldFileBase);
        if (index < files.size())
        {
            editors::ShowBitfieldDefinitionEditor(hwnd_, files[index].path);
        }
        return true;
    }
    switch (command_id)
    {
    case cmd::kToolsBitfieldDefinitions:
        editors::ShowBitfieldDefinitionEditor(hwnd_, std::wstring());
        return true;
    case cmd::kToolsKeyHandles:
        if (IsWindow(key_handles_window_))
        {
            ShowWindow(key_handles_window_, SW_RESTORE);
            SetForegroundWindow(key_handles_window_);
            return true;
        }
        key_handles_window_ = ShowKeyHandlesWindow(hwnd_, [this](const std::wstring& path, bool new_tab) {
            if (IsIconic(hwnd_))
            {
                ShowWindow(hwnd_, SW_RESTORE);
            }
            if (new_tab)
            {
                OpenLocalRegistryTab();
            }
            else
            {
                ActivateLocalRegistryTab();
            }
            ApplyViewVisibility();
            UpdateStatus();
            if (!SelectTreePath(path))
            {
                ui::ShowError(key_handles_window_, L"The key couldn't be found.");
                return;
            }
            SetForegroundWindow(hwnd_);
        });
        return true;
    case cmd::kEditDecodeValue:
        {
            const RegistryNode* node = browse_.current_node();
            if (!node)
            {
                return true;
            }
            std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
            if (selected_rows.size() != 1 || selected_rows.front().kind != rowkind::kValue)
            {
                return true;
            }
            ValueEntry entry;
            if (!RegistryStore::QueryValue(*node, selected_rows.front().extra, &entry))
            {
                ui::ShowError(hwnd_, L"Failed to read value.");
                return true;
            }
            editors::DecodeRequest request;
            request.value_name = entry.name;
            request.key_path = registry_path::Build(*node);
            request.type = entry.type;
            request.data = std::move(entry.data);
            editors::ShowValueDecoder(hwnd_, request);
            return true;
        }
    default:
        return false;
    }
}

} // namespace regkit
