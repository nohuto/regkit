// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/undo_stack.h"

#include <iterator>

namespace regkit::changes
{

void UndoStack::Push(UndoOperation operation)
{
    undo_.push_back(std::move(operation));
    redo_.clear();
}

void UndoStack::Push(std::vector<UndoOperation> steps)
{
    if (steps.size() == 1)
    {
        Push(std::move(steps.front()));
        return;
    }
    if (!steps.empty())
    {
        UndoOperation group;
        group.type = UndoOperation::Type::kGroup;
        group.steps = std::move(steps);
        Push(std::move(group));
    }
}

void UndoStack::CompletePartial(UndoOperation group, size_t replayed, bool redo, bool drop_failed)
{
    // undo replays from the back, redo from the front
    auto& steps = group.steps;
    const auto first = redo ? steps.begin() : steps.end() - static_cast<std::ptrdiff_t>(replayed);
    const auto last = first + static_cast<std::ptrdiff_t>(replayed);
    UndoOperation done;
    done.type = UndoOperation::Type::kGroup;
    done.steps.assign(std::make_move_iterator(first), std::make_move_iterator(last));
    steps.erase(first, last);
    // next step replay would have reached
    if (drop_failed && !steps.empty())
    {
        steps.erase(redo ? steps.begin() : steps.end() - 1);
    }
    if (!done.steps.empty())
    {
        (redo ? undo_ : redo_).push_back(std::move(done));
    }
    if (!steps.empty())
    {
        (redo ? redo_ : undo_).push_back(std::move(group));
    }
}

void UndoStack::ClearRedo()
{
    redo_.clear();
}

std::optional<UndoOperation> UndoStack::TakeUndo()
{
    if (undo_.empty())
    {
        return std::nullopt;
    }
    UndoOperation operation = std::move(undo_.back());
    undo_.pop_back();
    return operation;
}

std::optional<UndoOperation> UndoStack::TakeRedo()
{
    if (redo_.empty())
    {
        return std::nullopt;
    }
    UndoOperation operation = std::move(redo_.back());
    redo_.pop_back();
    return operation;
}

void UndoStack::CompleteUndo(UndoOperation operation)
{
    redo_.push_back(std::move(operation));
}

void UndoStack::CompleteRedo(UndoOperation operation)
{
    undo_.push_back(std::move(operation));
}

bool UndoStack::CanUndo() const noexcept
{
    return !undo_.empty();
}

bool UndoStack::CanRedo() const noexcept
{
    return !redo_.empty();
}

}
