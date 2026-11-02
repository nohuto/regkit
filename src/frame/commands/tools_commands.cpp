// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"

#include "dialogs/bitfield_definition_editor.h"
#include "dialogs/decoder_dialog.h"
#include "frame/tools/convert_dialog.h"
#include "frame/tools/key_handles_window.h"
#include "regfile/registry_transfer.h"
#include "registry/registry_path.h"
#include "win32/translation.h"

#include <win32ui/window.h>

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
    case cmd::kToolsConvertFile:
        {
            ConvertSettings settings;
            workspace::DialogState state = LoadDialogState();
            workspace::DialogFields last(&state, false);
            ConvertDialogFields(last, &settings);
            if (settings.source == ConvertSource::kRegistry && browse_.current_node())
            {
                settings.key_path = registry_path::Build(*browse_.current_node());
                settings.output_path = DefaultExportPath(settings.key_path, regfile::FormatExtension(settings.format));
            }
            const workspace::DialogState before = state;
            ShowConvertDialog(hwnd_, &settings);
            workspace::DialogFields store(&state, true);
            ConvertDialogFields(store, &settings);
            if (state != before)
            {
                workspace::SaveDialogState(state);
            }
        }
        return true;
    case cmd::kToolsKeyHandles:
        if (IsWindow(key_handles_window_))
        {
            win32ui::BringToFront(key_handles_window_);
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
                ui::ShowError(key_handles_window_, util::Tr(L"The key couldn't be found."));
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
            RegistryValue entry;
            if (!RegistryStore::QueryValue(*node, selected_rows.front().extra, &entry))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
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

}
