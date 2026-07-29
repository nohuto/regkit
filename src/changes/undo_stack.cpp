// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

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
