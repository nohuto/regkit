// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/undo_stack.h"

namespace regkit::changes {

void UndoStack::Push(UndoOperation operation) {
  undo_.push_back(std::move(operation));
  redo_.clear();
}

void UndoStack::ClearRedo() {
  redo_.clear();
}

std::optional<UndoOperation> UndoStack::TakeUndo() {
  if (undo_.empty()) {
    return std::nullopt;
  }
  UndoOperation operation = std::move(undo_.back());
  undo_.pop_back();
  return operation;
}

std::optional<UndoOperation> UndoStack::TakeRedo() {
  if (redo_.empty()) {
    return std::nullopt;
  }
  UndoOperation operation = std::move(redo_.back());
  redo_.pop_back();
  return operation;
}

void UndoStack::CompleteUndo(UndoOperation operation) {
  redo_.push_back(std::move(operation));
}

void UndoStack::CompleteRedo(UndoOperation operation) {
  undo_.push_back(std::move(operation));
}

bool UndoStack::CanUndo() const noexcept {
  return !undo_.empty();
}

bool UndoStack::CanRedo() const noexcept {
  return !redo_.empty();
}

} // namespace regkit::changes
