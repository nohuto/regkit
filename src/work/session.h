// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdint>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
#include <thread>
#include <utility>

namespace regkit::work {

class MoveOnly {
public:
  MoveOnly() = default;
  MoveOnly(MoveOnly&&) noexcept = default;
  MoveOnly& operator=(MoveOnly&&) noexcept = default;
  MoveOnly(const MoveOnly&) = delete;
  MoveOnly& operator=(const MoveOnly&) = delete;
};

class Session {
public:
  using Task = std::function<void(uint64_t, std::atomic_bool&)>;

  Session() = default;
  ~Session();
  Session(const Session&) = delete;
  Session& operator=(const Session&) = delete;

  uint64_t Start(Task task);
  bool StartIfIdle(Task task, uint64_t* generation = nullptr);
  void Cancel() noexcept;
  void CancelAndJoin() noexcept;
  void Join() noexcept;

  bool IsCurrent(uint64_t generation) const noexcept;
  bool running() const noexcept;
  uint64_t generation() const noexcept;

private:
  uint64_t StartPrepared(Task task);

  std::thread thread_;
  std::atomic_bool cancel_{false};
  std::atomic_bool running_{false};
  std::atomic<uint64_t> generation_{0};
};

template <typename Task>
class LatestTask {
public:
  using Processor =
      std::function<void(std::unique_ptr<Task>, const std::atomic_bool&)>;

  LatestTask() = default;
  ~LatestTask() {
    Stop();
  }
  LatestTask(const LatestTask&) = delete;
  LatestTask& operator=(const LatestTask&) = delete;

  void Start(Processor processor) {
    Stop();
    {
      std::lock_guard<std::mutex> lock(mutex_);
      processor_ = std::move(processor);
      stopping_.store(false);
    }
    thread_ = std::thread([this]() { Run(); });
  }

  void Submit(std::unique_ptr<Task> task) {
    if (!task) {
      return;
    }
    {
      std::lock_guard<std::mutex> lock(mutex_);
      if (stopping_.load()) {
        return;
      }
      pending_ = std::move(task);
    }
    ready_.notify_one();
  }

  void Stop() noexcept {
    stopping_.store(true);
    ready_.notify_one();
    if (thread_.joinable()) {
      thread_.join();
    }
    std::lock_guard<std::mutex> lock(mutex_);
    pending_.reset();
    processor_ = {};
  }

  bool running() const noexcept {
    return thread_.joinable() && !stopping_.load();
  }

private:
  void Run() {
    for (;;) {
      std::unique_ptr<Task> task;
      Processor processor;
      {
        std::unique_lock<std::mutex> lock(mutex_);
        ready_.wait(lock, [this]() {
          return stopping_.load() || pending_ != nullptr;
        });
        if (stopping_.load()) {
          return;
        }
        task = std::move(pending_);
        processor = processor_;
      }
      if (processor && task) {
        processor(std::move(task), stopping_);
      }
    }
  }

  mutable std::mutex mutex_;
  std::condition_variable ready_;
  std::thread thread_;
  std::atomic_bool stopping_{true};
  std::unique_ptr<Task> pending_;
  Processor processor_;
};

template <typename Task>
class DebouncedTask {
public:
  using Handler = std::function<void(Task)>;

  DebouncedTask() = default;
  ~DebouncedTask() {
    Stop();
  }
  DebouncedTask(const DebouncedTask&) = delete;
  DebouncedTask& operator=(const DebouncedTask&) = delete;

  void Start(std::chrono::milliseconds delay, Handler handler) {
    Stop();
    {
      std::lock_guard<std::mutex> lock(mutex_);
      delay_ = delay;
      handler_ = std::move(handler);
      stopping_ = false;
      revision_ = 0;
    }
    thread_ = std::thread([this]() { Run(); });
  }

  void Submit(Task task) {
    {
      std::lock_guard<std::mutex> lock(mutex_);
      if (stopping_) {
        return;
      }
      pending_ = std::move(task);
      ++revision_;
    }
    changed_.notify_one();
  }

  void Stop() noexcept {
    {
      std::lock_guard<std::mutex> lock(mutex_);
      stopping_ = true;
    }
    changed_.notify_one();
    if (thread_.joinable()) {
      thread_.join();
    }
    std::lock_guard<std::mutex> lock(mutex_);
    pending_.reset();
    handler_ = {};
  }

  bool running() const noexcept {
    return thread_.joinable();
  }

private:
  void Run() {
    std::unique_lock<std::mutex> lock(mutex_);
    for (;;) {
      changed_.wait(lock,
                    [this]() { return stopping_ || pending_.has_value(); });
      if (stopping_) {
        return;
      }

      uint64_t observed = revision_;
      auto deadline = std::chrono::steady_clock::now() + delay_;
      while (!stopping_ &&
             changed_.wait_until(lock, deadline, [this, observed]() {
               return stopping_ || revision_ != observed;
             })) {
        if (stopping_) {
          return;
        }
        observed = revision_;
        deadline = std::chrono::steady_clock::now() + delay_;
      }
      if (stopping_) {
        return;
      }

      Task task = std::move(*pending_);
      pending_.reset();
      Handler handler = handler_;
      lock.unlock();
      if (handler) {
        handler(std::move(task));
      }
      lock.lock();
    }
  }

  mutable std::mutex mutex_;
  std::condition_variable changed_;
  std::thread thread_;
  bool stopping_ = true;
  uint64_t revision_ = 0;
  std::chrono::milliseconds delay_{0};
  std::optional<Task> pending_;
  Handler handler_;
};

} // namespace regkit::work
