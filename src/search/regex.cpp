// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/regex.h"
#include "win32/translation.h"

#include <pcre2.h>

#include <algorithm>
#include <memory>
#include <utility>

namespace regkit::search::regex
{

namespace
{

PCRE2_SPTR16 Units(const wchar_t* text)
{
    static const wchar_t kEmpty[] = L"";
    return reinterpret_cast<PCRE2_SPTR16>(text ? text : kEmpty);
}

std::wstring ErrorMessage(int code)
{
    wchar_t buffer[256] = {};
    const int length = pcre2_get_error_message_16(code, reinterpret_cast<PCRE2_UCHAR16*>(buffer), std::size(buffer));
    return length > 0 ? std::wstring(buffer, static_cast<size_t>(length)) : std::wstring();
}

Status MapError(int code)
{
    switch (code)
    {
    case PCRE2_ERROR_NOMATCH:
    case PCRE2_ERROR_PARTIAL:
        return Status::kNoMatch;
    case PCRE2_ERROR_MATCHLIMIT:
    case PCRE2_ERROR_DEPTHLIMIT:
    case PCRE2_ERROR_HEAPLIMIT:
    case PCRE2_ERROR_NOMEMORY:
        return Status::kLimit;
    case PCRE2_ERROR_UTF16_ERR1:
    case PCRE2_ERROR_UTF16_ERR2:
    case PCRE2_ERROR_UTF16_ERR3:
    case PCRE2_ERROR_BADUTFOFFSET:
        return Status::kInvalidSubject;
    default:
        break;
    }
    return Status::kFailed;
}

}

class Pattern
{
  public:
    Pattern(pcre2_code_16* code, bool whole) noexcept
        : code_(code), whole_(whole)
    {
    }
    ~Pattern()
    {
        pcre2_code_free_16(code_);
    }
    Pattern(const Pattern&) = delete;
    Pattern& operator=(const Pattern&) = delete;

    pcre2_code_16* code() const noexcept
    {
        return code_;
    }
    bool whole() const noexcept
    {
        return whole_;
    }

  private:
    pcre2_code_16* code_ = nullptr;
    bool whole_ = false;
};

struct SessionState
{
    pcre2_match_data_16* data = nullptr;
    pcre2_match_context_16* context = nullptr;

    ~SessionState()
    {
        pcre2_match_data_free_16(data);
        pcre2_match_context_free_16(context);
    }
};

PatternRef Compile(const std::wstring& pattern, const Options& options, Error* error)
{
    if (error)
    {
        *error = Error();
    }
    if (pattern.empty())
    {
        return nullptr;
    }
    pcre2_compile_context_16* compile_context = pcre2_compile_context_create_16(nullptr);
    if (!compile_context)
    {
        return nullptr;
    }
    pcre2_set_max_pattern_length_16(compile_context, kMaxPatternLength);
    pcre2_set_parens_nest_limit_16(compile_context, kParensNestLimit);

    uint32_t compile_options = PCRE2_UTF | PCRE2_UCP | PCRE2_MATCH_INVALID_UTF | PCRE2_NEVER_BACKSLASH_C;
    if (options.ignore_case)
    {
        compile_options |= PCRE2_CASELESS;
    }
    if (options.whole)
    {
        compile_options |= PCRE2_ANCHORED | PCRE2_ENDANCHORED;
    }

    int code = 0;
    PCRE2_SIZE offset = 0;
    std::unique_ptr<pcre2_code_16, decltype(&pcre2_code_free_16)> compiled(
        pcre2_compile_16(Units(pattern.c_str()), pattern.size(), compile_options, &code, &offset, compile_context),
        pcre2_code_free_16
    );
    pcre2_compile_context_free_16(compile_context);
    if (!compiled)
    {
        if (error)
        {
            error->code = code;
            error->offset = offset;
            error->message = ErrorMessage(code);
        }
        return nullptr;
    }
    auto result = std::make_shared<const Pattern>(compiled.get(), options.whole);
    compiled.release();
    return result;
}

Session::Session() noexcept = default;

Session::Session(PatternRef pattern)
    : pattern_(std::move(pattern))
{
    if (!pattern_)
    {
        return;
    }
    auto state = std::make_unique<SessionState>();
    state->data = pcre2_match_data_create_from_pattern_16(pattern_->code(), nullptr);
    state->context = pcre2_match_context_create_16(nullptr);
    if (!state->data || !state->context)
    {
        pattern_.reset();
        return;
    }
    pcre2_set_match_limit_16(state->context, kMatchLimit);
    pcre2_set_depth_limit_16(state->context, kDepthLimit);
    pcre2_set_heap_limit_16(state->context, static_cast<uint32_t>(kHeapLimitKib));
    state_ = std::move(state);
}

Session::~Session() = default;
Session::Session(Session&&) noexcept = default;
Session& Session::operator=(Session&&) noexcept = default;

bool Session::valid() const noexcept
{
    return pattern_ && state_;
}

Found Session::Find(std::wstring_view subject) const
{
    Found found;
    if (!valid())
    {
        found.status = Status::kFailed;
        return found;
    }
    const int rc =
        pcre2_match_16(pattern_->code(), Units(subject.data()), subject.size(), 0, 0, state_->data, state_->context);
    if (rc < 0)
    {
        found.status = MapError(rc);
        return found;
    }
    const PCRE2_SIZE* output = pcre2_get_ovector_pointer_16(state_->data);
    found.status = Status::kMatch;
    found.start = output[0];
    found.length = output[1] - output[0];
    return found;
}

Status Session::Replace(std::wstring_view subject, const std::wstring& replacement, std::wstring* out, size_t* replacements) const
{
    if (replacements)
    {
        *replacements = 0;
    }
    if (!valid() || !out)
    {
        return Status::kFailed;
    }
    uint32_t options = PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_UNSET_EMPTY;
    if (!pattern_->whole())
    {
        options |= PCRE2_SUBSTITUTE_GLOBAL;
    }
    std::wstring buffer(std::min(subject.size() + replacement.size() + 32, kMaxReplaceLength + 1), L'\0');
    for (int attempt = 0; attempt < 2; ++attempt)
    {
        PCRE2_SIZE length = buffer.size();
        const int rc = pcre2_substitute_16(
            pattern_->code(),
            Units(subject.data()),
            subject.size(),
            0,
            options,
            state_->data,
            state_->context,
            Units(replacement.c_str()),
            replacement.size(),
            reinterpret_cast<PCRE2_UCHAR16*>(buffer.data()),
            &length
        );
        if (rc >= 0)
        {
            if (rc == 0)
            {
                return Status::kNoMatch;
            }
            if (length > kMaxReplaceLength)
            {
                return Status::kLimit;
            }
            buffer.resize(length);
            *out = std::move(buffer);
            if (replacements)
            {
                *replacements = static_cast<size_t>(rc);
            }
            return Status::kMatch;
        }
        if (rc != PCRE2_ERROR_NOMEMORY || attempt != 0)
        {
            return MapError(rc);
        }
        if (length > kMaxReplaceLength)
        {
            return Status::kLimit;
        }
        buffer.assign(length, L'\0');
    }
    return Status::kFailed;
}

std::wstring StatusText(Status status)
{
    switch (status)
    {
    case Status::kCancelled:
        return util::Tr(L"The search was cancelled.");
    case Status::kLimit:
        return util::Tr(L"The pattern needed too many steps or too much memory.");
    case Status::kInvalidSubject:
        return util::Tr(L"The text isn't valid UTF-16.");
    case Status::kFailed:
        return util::Tr(L"The pattern couldn't be applied.");
    default:
        break;
    }
    return std::wstring();
}

std::wstring ErrorText(const Error& error)
{
    if (error.message.empty())
    {
        return util::Tr(L"The find text isn't a valid regular expression.");
    }
    return util::TrDetail(L"The find text isn't a valid regular expression.", util::TrLabel(L"Position", std::to_wstring(error.offset + 1)) + L"\n" + error.message);
}

}
