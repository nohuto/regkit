// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "work/session.h"

namespace regkit::work {

Session::~Session() {
  CancelAndJoin();
}

uint64_t Session::Start(Task task) {
  CancelAndJoin();
  return StartPrepared(std::move(task));
}

bool Session::StartIfIdle(Task task, uint64_t* generation) {
  if (running_.load()) {
    return false;
  }
  Join();
  const uint64_t started = StartPrepared(std::move(task));
  if (generation) {
    *generation = started;
  }
  return true;
}

void Session::Cancel() noexcept {
  cancel_.store(true);
  generation_.fetch_add(1);
}

void Session::CancelAndJoin() noexcept {
  Cancel();
  Join();
}

void Session::Join() noexcept {
  if (thread_.joinable()) {
    thread_.join();
  }
  running_.store(false);
}

bool Session::IsCurrent(uint64_t generation) const noexcept {
  return generation_.load() == generation && !cancel_.load();
}

bool Session::running() const noexcept {
  return running_.load();
}

uint64_t Session::generation() const noexcept {
  return generation_.load();
}

uint64_t Session::StartPrepared(Task task) {
  cancel_.store(false);
  const uint64_t generation = generation_.fetch_add(1) + 1;
  running_.store(true);
  thread_ = std::thread(
      [this, generation, task = std::move(task)]() mutable {
        if (task) {
          task(generation, cancel_);
        }
        if (generation_.load() == generation) {
          running_.store(false);
        }
      });
  return generation;
}

} // namespace regkit::work
