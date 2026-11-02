// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/bitfield_definition.h"

#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/json/reader.h>
#include <win32ui/json/writer.h>
#include <win32ui/system/paths.h>

#include <algorithm>

namespace regkit::editors::bitfield
{

namespace
{

constexpr wchar_t kFormat[] = L"regkit-bitfield";
constexpr wchar_t kExtension[] = L".regkit-bitfield.jsonc";
constexpr int kMaxDepth = 8;

enum Member : unsigned
{
    kFormatMember = 1u << 0,
    kNameMember = 1u << 1,
    kCommentMember = 1u << 2,
    kDefinitionsMember = 1u << 3,
    kValueNameMember = 1u << 4,
    kKeyPathsMember = 1u << 5,
    kBitWidthMember = 1u << 6,
    kByteOffsetMember = 1u << 7,
    kFieldsMember = 1u << 8,
    kBitsMember = 1u << 9,
    kMeaningMember = 1u << 10,
    kStatesMember = 1u << 11,
    kValueMember = 1u << 12,
};

template <typename Read>
bool ReadMembers(win32ui::JsonReader& reader, unsigned* seen, Read&& read)
{
    *seen = 0;
    return reader.Object([&](const std::wstring& name) {
        unsigned flag = 0;
        if (!read(name, &flag))
        {
            return false;
        }
        if (flag == 0)
        {
            return reader.Fail(util::Tr(L"An object contains an unknown member."));
        }
        if (*seen & flag)
        {
            return reader.Fail(util::Tr(L"An object contains a duplicate member."));
        }
        *seen |= flag;
        return true;
    });
}

bool Require(win32ui::JsonReader& reader, unsigned seen, unsigned required, const wchar_t* message)
{
    return (seen & required) == required || reader.Fail(message);
}

bool ReadState(win32ui::JsonReader& reader, State* state)
{
    unsigned seen = 0;
    return ReadMembers(reader, &seen, [&](const std::wstring& name, unsigned* flag) {
               if (name == L"value")
               {
                   *flag = kValueMember;
                   return reader.Unsigned(&state->value);
               }
               if (name == L"name")
               {
                   *flag = kNameMember;
                   return reader.String(&state->name, kMaxNameLength);
               }
               if (name == L"meaning")
               {
                   *flag = kMeaningMember;
                   return reader.String(&state->meaning, kMaxMeaningLength);
               }
               return true;
           }) &&
           Require(reader, seen, kValueMember | kNameMember, util::Tr(L"A state is missing its value or name."));
}

bool ReadField(win32ui::JsonReader& reader, Field* field)
{
    unsigned seen = 0;
    return ReadMembers(reader, &seen, [&](const std::wstring& name, unsigned* flag) {
               if (name == L"name")
               {
                   *flag = kNameMember;
                   return reader.String(&field->name, kMaxNameLength);
               }
               if (name == L"bits")
               {
                   *flag = kBitsMember;
                   return reader.Array([&] {
                       uint64_t bit = 0;
                       if (!reader.Unsigned(&bit))
                       {
                           return false;
                       }
                       if (bit >= 64)
                       {
                           return reader.Fail(util::Tr(L"A field uses a bit outside the declared width."));
                       }
                       if (!field->bits.empty() && bit <= field->bits.back())
                       {
                           return reader.Fail(util::Tr(L"Field bits must be unique and in ascending order."));
                       }
                       field->bits.push_back(static_cast<unsigned>(bit));
                       return true;
                   }) && (!field->bits.empty() || reader.Fail(util::Tr(L"A field lists no bits.")));
               }
               if (name == L"meaning")
               {
                   *flag = kMeaningMember;
                   return reader.String(&field->meaning, kMaxMeaningLength);
               }
               if (name == L"states")
               {
                   *flag = kStatesMember;
                   return reader.Array([&] {
                       return (field->states.size() < kMaxStates ||
                               reader.Fail(util::Tr(L"A field lists more than 256 states."))) &&
                              ReadState(reader, &field->states.emplace_back());
                   });
               }
               return true;
           }) &&
           Require(reader, seen, kNameMember | kBitsMember, util::Tr(L"A field is missing its name or bits."));
}

bool ReadDefinition(win32ui::JsonReader& reader, Definition* definition)
{
    unsigned seen = 0;
    return ReadMembers(reader, &seen, [&](const std::wstring& name, unsigned* flag) {
               uint64_t number = 0;
               if (name == L"name")
               {
                   *flag = kNameMember;
                   return reader.String(&definition->name, kMaxNameLength);
               }
               if (name == L"value_name")
               {
                   *flag = kValueNameMember;
                   return reader.String(&definition->value_name, kMaxNameLength);
               }
               if (name == L"key_paths")
               {
                   *flag = kKeyPathsMember;
                   return reader.Array([&] {
                       return (definition->key_paths.size() < kMaxPaths ||
                               reader.Fail(util::Tr(L"A definition lists too many key paths."))) &&
                              reader.String(&definition->key_paths.emplace_back(), kMaxPathLength);
                   });
               }
               if (name == L"bit_width")
               {
                   *flag = kBitWidthMember;
                   definition->bit_width = 0;
                   return reader.Unsigned(&number) &&
                          ((ValidWidth(static_cast<unsigned>(number)) && number <= 64) ||
                           reader.Fail(util::Tr(L"The bit width must be 8, 16, 32, or 64."))) &&
                          (definition->bit_width = static_cast<unsigned>(number), true);
               }
               if (name == L"byte_offset")
               {
                   *flag = kByteOffsetMember;
                   return reader.Unsigned(&number) &&
                          (number <= kMaxByteOffset || reader.Fail(util::Tr(L"The byte offset is out of range."))) &&
                          (definition->byte_offset = static_cast<unsigned>(number), true);
               }
               if (name == L"comment")
               {
                   *flag = kCommentMember;
                   return reader.String(&definition->comment, kMaxCommentLength);
               }
               if (name == L"fields")
               {
                   *flag = kFieldsMember;
                   return reader.Array([&] {
                       return (definition->fields.size() < 64 ||
                               reader.Fail(util::Tr(L"A definition contains more than 64 fields."))) &&
                              ReadField(reader, &definition->fields.emplace_back());
                   });
               }
               return true;
           }) &&
           Require(reader, seen, kValueNameMember | kBitWidthMember, util::Tr(L"A definition is missing its value name or bit width."));
}

bool ReadFile(win32ui::JsonReader& reader, DefinitionFile* file)
{
    unsigned seen = 0;
    return ReadMembers(
               reader,
               &seen,
               [&](const std::wstring& name, unsigned* flag) {
                   if (name == L"format")
                   {
                       *flag = kFormatMember;
                       std::wstring format;
                       return reader.String(&format, kMaxNameLength) &&
                              (format == kFormat || reader.Fail(util::Tr(L"The file isn't a RegKit bitfield definition.")));
                   }
                   if (name == L"name")
                   {
                       *flag = kNameMember;
                       return reader.String(&file->name, kMaxNameLength);
                   }
                   if (name == L"comment")
                   {
                       *flag = kCommentMember;
                       return reader.String(&file->comment, kMaxCommentLength);
                   }
                   if (name == L"definitions")
                   {
                       *flag = kDefinitionsMember;
                       return reader.Array([&] { return ReadDefinition(reader, &file->definitions.emplace_back()); }) &&
                              (!file->definitions.empty() || reader.Fail(util::Tr(L"The file contains no definitions.")));
                   }
                   return true;
               }
           ) &&
           reader.End() &&
           Require(reader, seen, kFormatMember | kDefinitionsMember, util::Tr(L"The file is missing a required member."));
}

void AppendMember(std::wstring* out, const wchar_t* indent, const wchar_t* name, const std::wstring& text)
{
    out->append(indent).append(L"\"").append(name).append(L"\": ");
    win32ui::AppendJsonString(out, text);
}

std::wstring BundleDirectory()
{
    const std::wstring module_dir = util::ModuleDirectory();
    if (module_dir.empty())
    {
        return L"";
    }
    return util::JoinPath(util::JoinPath(module_dir, L"assets"), L"bitfields");
}

std::vector<DefinitionFile> LoadBundledFiles()
{
    std::vector<DefinitionFile> files;
    const std::wstring directory = BundleDirectory();
    if (directory.empty())
    {
        return files;
    }
    util::ForEachFile(util::JoinPath(directory, L"*.regkit-bitfield.jsonc").c_str(), [&](const WIN32_FIND_DATAW& found) {
        DefinitionFile file;
        std::wstring error;
        // ignore invalid files without hiding remaining definitions
        if ((found.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 && Load(util::JoinPath(directory, found.cFileName), &file, &error))
        {
            files.push_back(std::move(file));
        }
        return true;
    });
    std::stable_sort(files.begin(), files.end(), [](const DefinitionFile& left, const DefinitionFile& right) {
        return util::CompareInsensitive(left.name, right.name) < 0;
    });
    return files;
}

}

uint64_t Field::Extract(uint64_t value) const
{
    uint64_t out = 0;
    for (size_t i = 0; i < bits.size(); ++i)
    {
        if ((value >> bits[i]) & 1ull)
        {
            out |= 1ull << i;
        }
    }
    return out;
}

uint64_t Field::Apply(uint64_t value, uint64_t field_value) const
{
    for (size_t i = 0; i < bits.size(); ++i)
    {
        const uint64_t bit = 1ull << bits[i];
        if ((field_value >> i) & 1ull)
        {
            value |= bit;
        }
        else
        {
            value &= ~bit;
        }
    }
    return value;
}

const State* Field::StateFor(uint64_t field_value) const
{
    for (const State& state : states)
    {
        if (state.value == field_value)
        {
            return &state;
        }
    }
    return nullptr;
}

int Definition::FieldIndexForBit(unsigned bit) const
{
    if (bit >= 64)
    {
        return -1;
    }
    const signed char index = owner[bit];
    if (index < 0 || static_cast<size_t>(index) >= fields.size())
    {
        return -1;
    }
    return index;
}

const Field* Definition::FieldForBit(unsigned bit) const
{
    const int index = FieldIndexForBit(bit);
    return index < 0 ? nullptr : &fields[static_cast<size_t>(index)];
}

bool Definition::MatchesPath(const std::wstring& key_path) const
{
    // no path filters means every matching value name applies
    if (key_paths.empty())
    {
        return true;
    }
    for (const std::wstring& fragment : key_paths)
    {
        if (util::ContainsInsensitive(key_path, fragment))
        {
            return true;
        }
    }
    return false;
}

std::wstring DisplayName(const Definition& definition)
{
    if (!definition.name.empty())
    {
        return definition.name;
    }
    if (!definition.value_name.empty())
    {
        return definition.value_name;
    }
    return util::Tr(L"Unnamed definition");
}

bool ValidWidth(unsigned bit_width)
{
    return bit_width == 8 || bit_width == 16 || bit_width == 32 || bit_width == 64;
}

uint64_t WidthMask(unsigned bit_width)
{
    if (bit_width >= 64)
    {
        return ~0ull;
    }
    return (1ull << bit_width) - 1ull;
}

void BuildLookup(Definition* definition)
{
    definition->owner.fill(-1);
    for (size_t i = 0; i < definition->fields.size(); ++i)
    {
        Field& field = definition->fields[i];
        field.mask = 0;
        for (const unsigned bit : field.bits)
        {
            if (bit < 64)
            {
                field.mask |= 1ull << bit;
                definition->owner[bit] = static_cast<signed char>(i);
            }
        }
    }
}

bool Validate(Definition* definition, std::wstring* error)
{
    const auto fail = [error](const wchar_t* message) {
        if (error)
        {
            *error = message;
        }
        return false;
    };
    if (!definition)
    {
        return false;
    }
    if (!ValidWidth(definition->bit_width))
    {
        return fail(util::Tr(L"The bit width must be 8, 16, 32, or 64."));
    }
    if (definition->byte_offset > kMaxByteOffset)
    {
        return fail(util::Tr(L"The byte offset is out of range."));
    }
    if (definition->fields.size() > 64)
    {
        return fail(util::Tr(L"A definition contains more than 64 fields."));
    }
    if (definition->name.size() > kMaxNameLength || definition->value_name.size() > kMaxNameLength)
    {
        return fail(util::Tr(L"A name is longer than the format allows."));
    }
    if (definition->comment.size() > kMaxCommentLength)
    {
        return fail(util::Tr(L"A comment is longer than the format allows."));
    }
    if (definition->key_paths.size() > kMaxPaths)
    {
        return fail(util::Tr(L"A definition lists too many key paths."));
    }
    for (const std::wstring& path : definition->key_paths)
    {
        if (path.empty() || path.size() > kMaxPathLength)
        {
            return fail(util::Tr(L"A key path is empty or longer than the format allows."));
        }
    }
    std::array<signed char, 64> claimed = {};
    claimed.fill(-1);
    for (size_t i = 0; i < definition->fields.size(); ++i)
    {
        Field& field = definition->fields[i];
        if (field.name.empty() || field.name.size() > kMaxNameLength)
        {
            return fail(util::Tr(L"Every field needs a name of at most 256 characters."));
        }
        if (field.meaning.size() > kMaxMeaningLength)
        {
            return fail(util::Tr(L"A field meaning is longer than the format allows."));
        }
        if (field.bits.empty())
        {
            return fail(util::Tr(L"Every field needs at least one bit."));
        }
        for (size_t j = 0; j < field.bits.size(); ++j)
        {
            if (field.bits[j] >= definition->bit_width)
            {
                return fail(util::Tr(L"A field uses a bit outside the declared width."));
            }
            if (j > 0 && field.bits[j] <= field.bits[j - 1])
            {
                return fail(util::Tr(L"Field bits must be unique and in ascending order."));
            }
            if (claimed[field.bits[j]] >= 0)
            {
                return fail(util::Tr(L"Two fields claim the same bit."));
            }
            claimed[field.bits[j]] = static_cast<signed char>(i);
        }
        const uint64_t limit = WidthMask(static_cast<unsigned>(field.bits.size()));
        if (field.states.size() > kMaxStates)
        {
            return fail(util::Tr(L"A field lists more than 256 states."));
        }
        for (size_t j = 0; j < field.states.size(); ++j)
        {
            const State& state = field.states[j];
            if (state.name.empty() || state.name.size() > kMaxNameLength)
            {
                return fail(util::Tr(L"Every state needs a name of at most 256 characters."));
            }
            if (state.meaning.size() > kMaxMeaningLength)
            {
                return fail(util::Tr(L"A state meaning is longer than the format allows."));
            }
            if (state.value > limit)
            {
                return fail(util::Tr(L"A state value doesn't fit the bits of its field."));
            }
            for (size_t k = 0; k < j; ++k)
            {
                if (field.states[k].value == state.value)
                {
                    return fail(util::Tr(L"Two states of one field share the same value."));
                }
            }
        }
        for (size_t j = 0; j < i; ++j)
        {
            if (util::EqualsInsensitive(definition->fields[j].name, field.name))
            {
                return fail(util::Tr(L"Two fields share the same name."));
            }
        }
    }
    std::stable_sort(definition->fields.begin(), definition->fields.end(), [](const Field& left, const Field& right) { return left.bits.front() < right.bits.front(); });
    for (Field& field : definition->fields)
    {
        std::stable_sort(field.states.begin(), field.states.end(), [](const State& left, const State& right) { return left.value < right.value; });
    }
    BuildLookup(definition);
    return true;
}

bool Validate(DefinitionFile* file, std::wstring* error)
{
    if (!file)
    {
        return false;
    }
    if (file->definitions.empty())
    {
        if (error)
        {
            *error = util::Tr(L"The file contains no definitions.");
        }
        return false;
    }
    if (file->name.size() > kMaxNameLength || file->comment.size() > kMaxCommentLength)
    {
        if (error)
        {
            *error = util::Tr(L"A file name or comment is longer than the format allows.");
        }
        return false;
    }
    for (Definition& definition : file->definitions)
    {
        if (!Validate(&definition, error))
        {
            return false;
        }
    }
    return true;
}

bool Parse(const std::vector<BYTE>& utf8, DefinitionFile* file, std::wstring* error)
{
    if (!file)
    {
        return false;
    }
    const BYTE* data = utf8.data();
    size_t size = utf8.size();
    if (size >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
    {
        data += 3;
        size -= 3;
    }
    const std::wstring text = win32ui::Utf8ToWide(std::string_view(reinterpret_cast<const char*>(data), size)).value_or(std::wstring());
    if (size > 0 && text.empty())
    {
        if (error)
        {
            *error = util::Tr(L"The definition file isn't valid UTF-8.");
        }
        return false;
    }
    DefinitionFile parsed;
    std::wstring message;
    win32ui::JsonReader reader(text, win32ui::JsonErrorText(&message), kMaxDepth);
    if (!ReadFile(reader, &parsed))
    {
        if (error)
        {
            *error = message.empty() ? util::Tr(L"The definition file isn't valid JSON.") : message;
        }
        return false;
    }
    if (!Validate(&parsed, error))
    {
        return false;
    }
    *file = std::move(parsed);
    return true;
}

bool Load(const std::wstring& path, DefinitionFile* file, std::wstring* error)
{
    std::vector<BYTE> bytes;
    if (win32ui::ReadFileBytes(path.c_str(), &bytes, kMaxFileBytes) != ERROR_SUCCESS || bytes.empty())
    {
        if (error)
        {
            *error = util::Tr(L"The definition file couldn't be read.");
        }
        return false;
    }
    if (!Parse(bytes, file, error))
    {
        return false;
    }
    file->path = path;
    return true;
}

std::wstring Serialize(const DefinitionFile& file)
{
    std::wstring out;
    out.append(L"{\n  \"format\": \"").append(kFormat).append(L"\",\n");
    if (!file.name.empty())
    {
        AppendMember(&out, L"  ", L"name", file.name);
        out.append(L",\n");
    }
    if (!file.comment.empty())
    {
        AppendMember(&out, L"  ", L"comment", file.comment);
        out.append(L",\n");
    }
    out.append(L"  \"definitions\": [\n");
    for (size_t d = 0; d < file.definitions.size(); ++d)
    {
        const Definition& definition = file.definitions[d];
        out.append(L"    {\n");
        if (!definition.name.empty())
        {
            AppendMember(&out, L"      ", L"name", definition.name);
            out.append(L",\n");
        }
        AppendMember(&out, L"      ", L"value_name", definition.value_name);
        if (!definition.key_paths.empty())
        {
            out.append(L",\n      \"key_paths\": [\n");
            for (size_t i = 0; i < definition.key_paths.size(); ++i)
            {
                out.append(L"        ");
                win32ui::AppendJsonString(&out, definition.key_paths[i]);
                out.append(i + 1 < definition.key_paths.size() ? L",\n" : L"\n");
            }
            out.append(L"      ]");
        }
        out.append(L",\n      \"bit_width\": ").append(std::to_wstring(definition.bit_width));
        if (definition.byte_offset != 0)
        {
            out.append(L",\n      \"byte_offset\": ").append(std::to_wstring(definition.byte_offset));
        }
        if (!definition.comment.empty())
        {
            out.append(L",\n");
            AppendMember(&out, L"      ", L"comment", definition.comment);
        }
        if (!definition.fields.empty())
        {
            out.append(L",\n      \"fields\": [\n");
            for (size_t i = 0; i < definition.fields.size(); ++i)
            {
                const Field& field = definition.fields[i];
                out.append(L"        {\n");
                AppendMember(&out, L"          ", L"name", field.name);
                out.append(L",\n          \"bits\": [");
                for (size_t j = 0; j < field.bits.size(); ++j)
                {
                    if (j > 0)
                    {
                        out.append(L", ");
                    }
                    out.append(std::to_wstring(field.bits[j]));
                }
                out.append(L"]");
                if (!field.meaning.empty())
                {
                    out.append(L",\n");
                    AppendMember(&out, L"          ", L"meaning", field.meaning);
                }
                if (!field.states.empty())
                {
                    out.append(L",\n          \"states\": [\n");
                    for (size_t j = 0; j < field.states.size(); ++j)
                    {
                        const State& state = field.states[j];
                        out.append(L"            { \"value\": ").append(std::to_wstring(state.value)).append(L", ");
                        AppendMember(&out, L"", L"name", state.name);
                        if (!state.meaning.empty())
                        {
                            out.append(L", ");
                            AppendMember(&out, L"", L"meaning", state.meaning);
                        }
                        out.append(L" }");
                        out.append(j + 1 < field.states.size() ? L",\n" : L"\n");
                    }
                    out.append(L"          ]");
                }
                out.append(L"\n        }");
                out.append(i + 1 < definition.fields.size() ? L",\n" : L"\n");
            }
            out.append(L"      ]");
        }
        out.append(L"\n    }");
        out.append(d + 1 < file.definitions.size() ? L",\n" : L"\n");
    }
    out.append(L"  ]\n}\n");
    return out;
}

bool Save(const std::wstring& path, const DefinitionFile& file, std::wstring* error)
{
    DefinitionFile copy = file;
    if (!Validate(&copy, error))
    {
        return false;
    }
    // formatting can exceed the Load size limit for otherwise valid data
    const std::wstring text = Serialize(copy);
    if (win32ui::WideToMultiByte(text, CP_UTF8).size() > kMaxFileBytes)
    {
        if (error)
        {
            *error = util::Tr(L"The file is larger than RegKit can load.");
        }
        return false;
    }
    if (win32ui::WriteTextFileAtomic(path.c_str(), text) != ERROR_SUCCESS)
    {
        if (error)
        {
            *error = util::Tr(L"The definition file couldn't be written.");
        }
        return false;
    }
    return true;
}

const std::vector<DefinitionFile>& BundledFiles()
{
    static const std::vector<DefinitionFile> files = LoadBundledFiles();
    return files;
}

std::vector<Definition> Matching(const std::wstring& key_path, const std::wstring& value_name)
{
    std::vector<Definition> matches;
    for (const DefinitionFile& file : BundledFiles())
    {
        for (const Definition& definition : file.definitions)
        {
            if (!util::EqualsInsensitive(definition.value_name, value_name))
            {
                continue;
            }
            if (!definition.MatchesPath(key_path))
            {
                continue;
            }
            matches.push_back(definition);
        }
    }
    return matches;
}

std::wstring SuggestedFileName(const std::wstring& value_name)
{
    const std::wstring safe = win32ui::SafeFileName(value_name);
    return (safe.empty() ? std::wstring(L"Default") : safe) + kExtension;
}

const wchar_t* FileFilter()
{
    return L"RegKit bitfield definitions (*.regkit-bitfield.jsonc)\0*.regkit-bitfield.jsonc\0JSON files "
           L"(*.jsonc;*.json)\0*.jsonc;*.json\0";
}

}
