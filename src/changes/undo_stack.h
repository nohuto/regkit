// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "changes/key_snapshot.h"

#include <win32ui/handle_owner.h>

#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace regkit::changes
{

struct UndoOperation
{
    enum class Type
    {
        kCreateKey,
        kDeleteKey,
        kRenameKey,
        kCreateValue,
        kDeleteValue,
        kModifyValue,
        kRenameValue,
        kReplaceKey,
        kGroup,
    };

    Type type = Type::kCreateKey;
    RegistryNode node;
    std::wstring name;
    std::wstring new_name;
    RegistryValue old_value;
    RegistryValue new_value;
    KeySnapshot key_snapshot;
    // kReplaceKey hive copies before and after, the files go away with the last handle
    std::shared_ptr<win32ui::UniqueFile> hive_before;
    std::shared_ptr<win32ui::UniqueFile> hive_after;
    std::vector<std::pair<RegistryNode, KeySnapshot>> volatile_keys;
    std::vector<UndoOperation> steps;
};

class UndoStack
{
  public:
    void Push(UndoOperation operation);
    void Push(std::vector<UndoOperation> steps);
    void CompletePartial(UndoOperation group, size_t replayed, bool redo, bool drop_failed = false);
    void ClearRedo();
    std::optional<UndoOperation> TakeUndo();
    std::optional<UndoOperation> TakeRedo();
    void CompleteUndo(UndoOperation operation);
    void CompleteRedo(UndoOperation operation);

    bool CanUndo() const noexcept;
    bool CanRedo() const noexcept;

  private:
    std::vector<UndoOperation> undo_;
    std::vector<UndoOperation> redo_;
};

}
