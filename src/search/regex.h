// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <atomic>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>

namespace regkit::search::regex
{

inline constexpr size_t kMaxPatternLength = 8192;
inline constexpr uint32_t kParensNestLimit = 100;
inline constexpr uint32_t kMatchLimit = 200000;
inline constexpr uint32_t kDepthLimit = 2000;
inline constexpr size_t kHeapLimitKib = 4096;
inline constexpr size_t kMaxReplaceLength = 4u * 1024u * 1024u;

enum class Status : uint8_t
{
    kMatch,
    kNoMatch,
    kCancelled,
    kLimit,
    kInvalidSubject,
    kFailed,
};

struct Options
{
    bool ignore_case = false;
    bool whole = false;
};

struct Error
{
    int code = 0;
    size_t offset = 0;
    std::wstring message;
};

struct Found
{
    Status status = Status::kNoMatch;
    size_t start = 0;
    size_t length = 0;
};

class Pattern;
using PatternRef = std::shared_ptr<const Pattern>;

PatternRef Compile(const std::wstring& pattern, const Options& options, Error* error);

struct SessionState;

// one session per thread
class Session
{
  public:
    Session() noexcept;
    explicit Session(PatternRef pattern);
    ~Session();
    Session(Session&&) noexcept;
    Session& operator=(Session&&) noexcept;

    bool valid() const noexcept;
    Found Find(std::wstring_view subject) const;
    Status Replace(std::wstring_view subject, const std::wstring& replacement, std::wstring* out, size_t* replacements) const;

  private:
    PatternRef pattern_;
    std::unique_ptr<SessionState> state_;
};

std::wstring StatusText(Status status);
std::wstring ErrorText(const Error& error);

}
