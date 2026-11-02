// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/resource_list.h"

#include "win32/text_transform.h"
#include "win32/translation.h"

#include <algorithm>
#include <array>
#include <cstring>
#include <span>
#include <utility>

namespace regkit::resource_list
{
namespace
{

// definitions from wdm.h
constexpr size_t kPartialHeader = 4;
constexpr size_t kRequirementSize = 32;
constexpr size_t kDeviceSpecificLimit = 256;
constexpr USHORT kInterruptLatched = 0x1;
constexpr USHORT kInterruptMessage = 0x2;
constexpr USHORT kMemoryLarge40 = 0x200;
constexpr USHORT kMemoryLarge48 = 0x400;
constexpr USHORT kMemoryLarge64 = 0x800;

struct Reader
{
    const BYTE* data = nullptr;
    size_t size = 0;
    bool ok = true;

    template <typename T>
    T At(size_t offset) noexcept
    {
        T value{};
        if (offset > size || sizeof(T) > size - offset)
        {
            ok = false;
            return value;
        }
        std::memcpy(&value, data + offset, sizeof(T));
        return value;
    }
    unsigned long long Affinity(size_t offset, size_t width) noexcept
    {
        return width == 8 ? At<unsigned long long>(offset) : At<ULONG>(offset);
    }
    std::wstring Bytes(size_t offset, size_t length, size_t limit = 0) const
    {
        const size_t start = std::min(offset, size);
        return win32ui::HexBytes(std::span(data + start, std::min(length, size - start)), L' ', true, limit);
    }
};

std::wstring Number(unsigned long long value)
{
    return std::to_wstring(value);
}

std::wstring NameOrNumber(std::span<const std::pair<int, const wchar_t*>> names, int value)
{
    for (const auto& [key, name] : names)
    {
        if (key == value)
        {
            return util::Tr(name);
        }
    }
    return win32ui::FormatHex(static_cast<unsigned long long>(static_cast<unsigned int>(value)), 2);
}

constexpr std::pair<int, const wchar_t*> kInterfaceTypes[] = {
    {-1, util::TrNoop(L"Undefined")},
    {0, util::TrNoop(L"Internal")},
    {1, L"ISA"},
    {2, L"EISA"},
    {3, L"MicroChannel"},
    {4, L"TurboChannel"},
    {5, L"PCI"},
    {6, L"VME"},
    {7, L"NuBus"},
    {8, L"PCMCIA"},
    {9, L"CBus"},
    {10, L"MPI"},
    {11, L"MPSA"},
    {12, util::TrNoop(L"Processor internal")},
    {13, util::TrNoop(L"Internal power bus")},
    {14, L"PnP ISA"},
    {15, L"PnP"},
    {16, L"VMCS"},
    {17, L"ACPI"},
};

constexpr std::pair<int, const wchar_t*> kResourceTypes[] = {
    {0, util::TrNoop(L"Null")},
    {0x80, util::TrNoop(L"Configuration data")},
    {0x81, util::TrNoop(L"Device private")},
    {0x82, util::TrNoop(L"PC card configuration")},
    {0x83, util::TrNoop(L"Multifunction card configuration")},
    {0x84, util::TrNoop(L"Connection")},
};

constexpr std::pair<int, const wchar_t*> kShareDispositions[] = {
    {0, util::TrNoop(L"Undetermined")},
    {1, util::TrNoop(L"Device exclusive")},
    {2, util::TrNoop(L"Driver exclusive")},
    {3, util::TrNoop(L"Shared")},
};

unsigned long long LargeLength(USHORT flags, ULONG length)
{
    const int shift = (flags & kMemoryLarge64) ? 32 : (flags & kMemoryLarge48) ? 16
                                                  : (flags & kMemoryLarge40)   ? 8
                                                                               : 0;
    return static_cast<unsigned long long>(length) << shift;
}

std::wstring Share(BYTE share)
{
    return NameOrNumber(kShareDispositions, share);
}

struct FlagSet
{
    // name applies when zero_mask bits are all clear, such as PORT_MEMORY or DMA_8
    const wchar_t* zero_name;
    USHORT zero_mask;
    std::initializer_list<std::pair<unsigned, const wchar_t*>> bits;
};

constexpr USHORT kDmaV3 = 0x80;

// wdm.h names without CM_RESOURCE_ prefix
const FlagSet kPortFlags = {L"PORT_MEMORY", 0x1, {{0x1, L"PORT_IO"}, {0x4, L"PORT_10_BIT_DECODE"}, {0x8, L"PORT_12_BIT_DECODE"}, {0x10, L"PORT_16_BIT_DECODE"}, {0x20, L"PORT_POSITIVE_DECODE"}, {0x40, L"PORT_PASSIVE_DECODE"}, {0x80, L"PORT_WINDOW_DECODE"}, {0x100, L"PORT_BAR"}}};
const FlagSet kMemoryFlags = {L"MEMORY_READ_WRITE", 0x3, {{0x1, L"MEMORY_READ_ONLY"}, {0x2, L"MEMORY_WRITE_ONLY"}, {0x4, L"MEMORY_PREFETCHABLE"}, {0x8, L"MEMORY_COMBINEDWRITE"}, {0x10, L"MEMORY_24"}, {0x20, L"MEMORY_CACHEABLE"}, {0x40, L"MEMORY_WINDOW_DECODE"}, {0x80, L"MEMORY_BAR"}, {0x100, L"MEMORY_COMPAT_FOR_INACCESSIBLE_RANGE"}, {kMemoryLarge40, L"MEMORY_LARGE_40"}, {kMemoryLarge48, L"MEMORY_LARGE_48"}, {kMemoryLarge64, L"MEMORY_LARGE_64"}}};
const FlagSet kInterruptFlags = {L"INTERRUPT_LEVEL_SENSITIVE", kInterruptLatched, {{kInterruptLatched, L"INTERRUPT_LATCHED"}, {kInterruptMessage, L"INTERRUPT_MESSAGE"}, {0x4, L"INTERRUPT_POLICY_INCLUDED"}, {0x10, L"INTERRUPT_SECONDARY_INTERRUPT"}, {0x20, L"INTERRUPT_WAKE_HINT"}}};
const FlagSet kDmaFlags = {L"DMA_8", 0x7, {{0x1, L"DMA_16"}, {0x2, L"DMA_32"}, {0x4, L"DMA_8_AND_16"}, {0x8, L"DMA_BUS_MASTER"}, {0x10, L"DMA_TYPE_A"}, {0x20, L"DMA_TYPE_B"}, {0x40, L"DMA_TYPE_F"}, {kDmaV3, L"DMA_V3"}}};

std::wstring Flags(USHORT flags, const FlagSet& set)
{
    std::wstring text = (flags & set.zero_mask) == 0 ? set.zero_name : L"";
    USHORT rest = flags;
    for (const auto& [bit, name] : set.bits)
    {
        if (flags & bit)
        {
            text.append(text.empty() ? L"" : L", ").append(name);
            rest = static_cast<USHORT>(rest & ~bit);
        }
    }
    if (rest)
    {
        text.append(text.empty() ? L"" : L", ").append(win32ui::FormatHex(rest, 0));
    }
    return text + L" (" + win32ui::FormatHex(flags, 4) + L")";
}

enum Kind
{
    kPort,
    kInterrupt,
    kMemory,
    kDma,
    kBusNumber,
    kDeviceSpecific,
    kOther,
    kKindCount,
};

constexpr const wchar_t* kKindTitles[kKindCount] = {
    util::TrNoop(L"Ports"),
    util::TrNoop(L"Interrupts"),
    util::TrNoop(L"Memory"),
    util::TrNoop(L"DMA"),
    util::TrNoop(L"Bus numbers"),
    util::TrNoop(L"Device specific data"),
    util::TrNoop(L"Other resources"),
};

Kind KindOf(BYTE type)
{
    switch (type)
    {
    case 1:
        return kPort;
    case 2:
        return kInterrupt;
    case 3:
    case 7:
        return kMemory;
    case 4:
        return kDma;
    case 5:
        return kDeviceSpecific;
    case 6:
        return kBusNumber;
    default:
        return kOther;
    }
}

struct Tables
{
    win32ui::TextTable header;
    std::array<win32ui::TextTable, kKindCount> kinds;
    const wchar_t* number_title = nullptr;
    std::wstring number;
    bool translated = false;

    void Add(Kind kind, std::initializer_list<const wchar_t*> columns, std::vector<std::wstring> cells)
    {
        win32ui::TextTable& table = kinds[kind];
        if (table.columns.empty())
        {
            table.title = util::Tr(kKindTitles[kind]);
            if (number_title)
            {
                table.columns.push_back(util::Tr(number_title));
            }
            for (const wchar_t* column : columns)
            {
                table.columns.push_back(util::Tr(column));
            }
        }
        if (number_title)
        {
            cells.insert(cells.begin(), number);
        }
        table.rows.push_back(std::move(cells));
    }
    std::vector<win32ui::TextTable> Take()
    {
        std::vector<win32ui::TextTable> tables;
        tables.reserve(kinds.size() + 1);
        tables.push_back(std::move(header));
        for (win32ui::TextTable& table : kinds)
        {
            if (!table.rows.empty())
            {
                tables.push_back(std::move(table));
            }
        }
        return tables;
    }
};

// descriptor byte count, including trailing device specific data
size_t DecodePartial(Reader& reader, size_t offset, size_t width, size_t descriptor, Tables& out)
{
    const BYTE type = reader.At<BYTE>(offset);
    const BYTE share = reader.At<BYTE>(offset + 1);
    const USHORT flags = reader.At<USHORT>(offset + 2);
    const size_t u = offset + kPartialHeader;
    const Kind kind = KindOf(type);
    switch (kind)
    {
    case kPort:
    case kMemory:
        {
            const ULONG length = reader.At<ULONG>(u + 8);
            out.Add(kind, {util::TrNoop(L"Start address"), util::TrNoop(L"Length"), util::TrNoop(L"Share"), util::TrNoop(L"Flags")}, {win32ui::FormatHex(reader.At<unsigned long long>(u), 16), type == 7 ? win32ui::FormatHex(LargeLength(flags, length), 16) : win32ui::FormatHex(length, 8), Share(share), Flags(flags, kind == kPort ? kPortFlags : kMemoryFlags)});
            break;
        }
    case kInterrupt:
        {
            // raw message interrupts store group/count where other interrupts store level/group
            const bool raw_message = (flags & kInterruptMessage) && !out.translated;
            std::wstring flag_text = Flags(flags, kInterruptFlags);
            if (raw_message)
            {
                flag_text.append(L", ").append(util::TrLabel(L"Messages", Number(reader.At<USHORT>(u + 2))));
            }
            out.Add(kind, {util::TrNoop(L"Vector"), util::TrNoop(L"Level"), util::TrNoop(L"Group"), util::TrNoop(L"Affinity"), util::TrNoop(L"Flags"), util::TrNoop(L"Share")}, {Number(reader.At<ULONG>(u + 4)), raw_message ? std::wstring() : Number(reader.At<USHORT>(u)), Number(reader.At<USHORT>(raw_message ? u : u + 2)), win32ui::FormatHex(reader.Affinity(u + 8, width), static_cast<int>(width * 2)), flag_text, Share(share)});
            break;
        }
    case kDma:
        out.Add(kind, {util::TrNoop(L"Channel"), util::TrNoop(L"Port / request line"), util::TrNoop(L"Share"), util::TrNoop(L"Flags")}, {Number(reader.At<ULONG>(u)), Number(reader.At<ULONG>(u + 4)), Share(share), Flags(flags, kDmaFlags)});
        break;
    case kBusNumber:
        out.Add(kind, {util::TrNoop(L"Start bus number"), util::TrNoop(L"Length"), util::TrNoop(L"Share")}, {Number(reader.At<ULONG>(u)), Number(reader.At<ULONG>(u + 4)), Share(share)});
        break;
    case kDeviceSpecific:
        {
            const ULONG length = reader.At<ULONG>(u);
            if (length > reader.size || offset + descriptor > reader.size - length)
            {
                reader.ok = false;
                return descriptor;
            }
            out.Add(kind, {util::TrNoop(L"Size"), util::TrNoop(L"Data")}, {Number(length), reader.Bytes(offset + descriptor, length, kDeviceSpecificLimit)});
            return descriptor + length;
        }
    default:
        out.Add(kind, {util::TrNoop(L"Type"), util::TrNoop(L"Share"), util::TrNoop(L"Flags"), util::TrNoop(L"Data")}, {NameOrNumber(kResourceTypes, type), Share(share), win32ui::FormatHex(flags, 4), type == 0x84 ? win32ui::FormatHex((static_cast<unsigned long long>(reader.At<ULONG>(u + 8)) << 32) | reader.At<ULONG>(u + 4), 16) : reader.Bytes(u, descriptor - kPartialHeader)});
        break;
    }
    return descriptor;
}

size_t DecodeFull(Reader& reader, size_t offset, size_t width, Tables& out)
{
    const size_t descriptor = kPartialHeader + std::max<size_t>(12, 8 + width);
    const ULONG count = reader.At<ULONG>(offset + 12);
    std::vector<std::wstring> row = {NameOrNumber(kInterfaceTypes, reader.At<LONG>(offset)), Number(reader.At<ULONG>(offset + 4)), Number(reader.At<USHORT>(offset + 8)) + L"." + Number(reader.At<USHORT>(offset + 10)), Number(count)};
    if (out.number_title)
    {
        row.insert(row.begin(), out.number);
    }
    out.header.rows.push_back(std::move(row));
    size_t position = offset + 16;
    for (ULONG partial = 0; partial < count && reader.ok; ++partial)
    {
        position += DecodePartial(reader, position, width, descriptor, out);
        reader.ok = reader.ok && position <= reader.size;
    }
    return position - offset;
}

void DecodeRequirement(Reader& reader, size_t offset, size_t width, Tables& out)
{
    static constexpr std::pair<int, const wchar_t*> kOptions[] = {
        {0, util::TrNoop(L"Required")},
        {1, util::TrNoop(L"Preferred")},
        {2, util::TrNoop(L"Default")},
        {8, util::TrNoop(L"Alternative")},
    };
    // IO_RESOURCE_PREFERRED, _DEFAULT, _ALTERNATIVE are bits, none of them is required
    const BYTE option_bits = reader.At<BYTE>(offset);
    std::wstring option;
    for (const auto& [bit, name] : kOptions)
    {
        if (option_bits & bit)
        {
            option.append(option.empty() ? L"" : L", ").append(util::Tr(name));
        }
    }
    if (option_bits & ~0x0B)
    {
        option.append(option.empty() ? L"" : L", ").append(win32ui::FormatHex(option_bits & ~0x0B, 2));
    }
    option = option.empty() ? std::wstring(util::Tr(kOptions[0].second)) : option;
    const BYTE type = reader.At<BYTE>(offset + 1);
    const std::wstring share = Share(reader.At<BYTE>(offset + 2));
    const USHORT flags = reader.At<USHORT>(offset + 4);
    const size_t u = offset + 8;
    const Kind kind = KindOf(type);
    switch (kind)
    {
    case kPort:
    case kMemory:
        {
            auto length = [&](size_t at) { return win32ui::FormatHex(type == 7 ? LargeLength(flags, reader.At<ULONG>(at)) : reader.At<ULONG>(at), 8); };
            out.Add(kind, {util::TrNoop(L"Option"), util::TrNoop(L"Length"), util::TrNoop(L"Alignment"), util::TrNoop(L"Minimum address"), util::TrNoop(L"Maximum address"), util::TrNoop(L"Share"), util::TrNoop(L"Flags")}, {option, length(u), length(u + 4), win32ui::FormatHex(reader.At<unsigned long long>(u + 8), 16), win32ui::FormatHex(reader.At<unsigned long long>(u + 16), 16), share, Flags(flags, kind == kPort ? kPortFlags : kMemoryFlags)});
            break;
        }
    case kInterrupt:
        out.Add(kind, {util::TrNoop(L"Option"), util::TrNoop(L"Minimum vector"), util::TrNoop(L"Maximum vector"), util::TrNoop(L"Group"), util::TrNoop(L"Targeted processors"), util::TrNoop(L"Flags"), util::TrNoop(L"Share")}, {option, Number(reader.At<ULONG>(u)), Number(reader.At<ULONG>(u + 4)), Number(reader.At<USHORT>(u + 10)), win32ui::FormatHex(reader.Affinity(u + 16, width), static_cast<int>(width * 2)), Flags(flags, kInterruptFlags), share});
        break;
    case kDma:
        {
            // DmaV3 layout is {RequestLine, Reserved, Channel, TransferWidth}, one channel
            const bool v3 = flags & kDmaV3;
            out.Add(kind, {util::TrNoop(L"Option"), util::TrNoop(L"Minimum channel"), util::TrNoop(L"Maximum channel"), util::TrNoop(L"Share"), util::TrNoop(L"Flags")}, {option, Number(reader.At<ULONG>(v3 ? u + 8 : u)), Number(reader.At<ULONG>(v3 ? u + 8 : u + 4)), share, Flags(flags, kDmaFlags)});
        }
        break;
    case kBusNumber:
        out.Add(kind, {util::TrNoop(L"Option"), util::TrNoop(L"Length"), util::TrNoop(L"Minimum bus number"), util::TrNoop(L"Maximum bus number"), util::TrNoop(L"Share")}, {option, Number(reader.At<ULONG>(u)), Number(reader.At<ULONG>(u + 4)), Number(reader.At<ULONG>(u + 8)), share});
        break;
    default:
        out.Add(kOther, {util::TrNoop(L"Option"), util::TrNoop(L"Type"), util::TrNoop(L"Share"), util::TrNoop(L"Flags"), util::TrNoop(L"Data")}, {option, NameOrNumber(kResourceTypes, type), share, win32ui::FormatHex(flags, 4), reader.Bytes(u, 24)});
        break;
    }
}

std::optional<std::vector<win32ui::TextTable>> DecodeLayout(DWORD type, const BYTE* data, size_t size, size_t width, bool translated, size_t* consumed)
{
    Reader reader{data, size};
    Tables out;
    out.translated = translated;
    size_t position = 0;
    if (type == REG_RESOURCE_REQUIREMENTS_LIST)
    {
        const ULONG list_size = reader.At<ULONG>(0);
        const ULONG lists = reader.At<ULONG>(28);
        out.header = {util::Tr(L"Requirements"), {util::Tr(L"Interface"), util::Tr(L"Bus number"), util::Tr(L"Slot number"), util::Tr(L"Alternative lists")}, {{NameOrNumber(kInterfaceTypes, reader.At<LONG>(4)), Number(reader.At<ULONG>(8)), Number(reader.At<ULONG>(12)), Number(lists)}}};
        out.number_title = lists > 1 ? util::TrNoop(L"List") : nullptr;
        position = 32;
        for (ULONG list = 0; list < lists && reader.ok; ++list)
        {
            out.number = Number(list + 1);
            const ULONG count = reader.At<ULONG>(position + 4);
            position += 8;
            for (ULONG item = 0; item < count && reader.ok; ++item, position += kRequirementSize)
            {
                DecodeRequirement(reader, position, width, out);
            }
        }
        reader.ok = reader.ok && position <= list_size && list_size <= size;
    }
    else
    {
        const bool single = type == REG_FULL_RESOURCE_DESCRIPTOR;
        const ULONG count = single ? 1 : reader.At<ULONG>(0);
        out.number_title = count > 1 ? util::TrNoop(L"Descriptor") : nullptr;
        out.header = {util::Tr(L"Descriptors"), {util::Tr(L"Interface"), util::Tr(L"Bus number"), util::Tr(L"Version"), util::Tr(L"Resources")}, {}};
        if (out.number_title)
        {
            out.header.columns.insert(out.header.columns.begin(), util::Tr(out.number_title));
        }
        position = single ? 0 : 4;
        for (ULONG index = 0; index < count && reader.ok; ++index)
        {
            out.number = Number(index + 1);
            position += DecodeFull(reader, position, width, out);
        }
    }
    *consumed = position;
    if (!reader.ok || position > size)
    {
        return std::nullopt;
    }
    return out.Take();
}

}

bool IsResourceType(DWORD type) noexcept
{
    return type == REG_RESOURCE_LIST || type == REG_FULL_RESOURCE_DESCRIPTOR || type == REG_RESOURCE_REQUIREMENTS_LIST;
}

std::optional<std::vector<win32ui::TextTable>> Decode(DWORD type, const BYTE* data, size_t size, bool translated)
{
    // prefer native layout, fall back if it doesnt account for the exact data size
    SYSTEM_INFO system = {};
    GetNativeSystemInfo(&system);
    const size_t native = system.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_INTEL ? 4 : 8;
    size_t consumed = 0;
    auto first = DecodeLayout(type, data, size, native, translated, &consumed);
    if (first && consumed == size)
    {
        return first;
    }
    size_t other_consumed = 0;
    auto other = DecodeLayout(type, data, size, native == 8 ? 4 : 8, translated, &other_consumed);
    return other && (other_consumed == size || !first) ? other : first;
}

}
