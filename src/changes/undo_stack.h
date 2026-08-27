// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "changes/key_snapshot.h"

#include <optional>
#include <string>
#include <vector>

namespace regkit::changes {

struct UndoOperation {
  enum class Type {
    kCreateKey,
    kDeleteKey,
    kRenameKey,
    kCreateValue,
    kDeleteValue,
    kModifyValue,
    kRenameValue,
  };

  Type type = Type::kCreateKey;
  RegistryNode node;
  std::wstring name;
  std::wstring new_name;
  ValueEntry old_value;
  ValueEntry new_value;
  KeySnapshot key_snapshot;
};

class UndoStack {
public:
  void Push(UndoOperation operation);
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

} // namespace regkit::changes
