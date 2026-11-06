// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/key_snapshot.h"

#include "registry/registry_path.h"
#include "win32/process_rights.h"

#include <algorithm>

namespace regkit::changes
{
namespace
{

constexpr SECURITY_INFORMATION kExactSecurity =
    OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION | LABEL_SECURITY_INFORMATION;

KeySnapshot Capture(const RegistryNode& node, SECURITY_INFORMATION parts, const std::atomic_bool* cancel)
{
    const auto cancelled = [cancel] { return cancel && cancel->load(); };
    if (cancelled())
    {
        return {.complete = false};
    }
    KeySnapshot snapshot;
    snapshot.name = registry_path::Leaf(node.subkey);
    snapshot.security_parts = parts;
    if (parts && !RegistryStore::ReadKeySecurity(node, &snapshot.security_parts, &snapshot.security) && !RegistryStore::IsVirtualRoot(node.root))
    {
        snapshot.complete = false;
    }
    const bool link = RegistryStore::ReadKeyLink(node, &snapshot.link_target);
    RegistryStore::KeyEnumResult result;
    result.want_options = true;
    bool reserved = false;
    std::vector<std::wstring> children;
    snapshot.complete = RegistryStore::EnumKeyStreaming(
                            node,
                            !link,
                            true,
                            !link,
                            &result,
                            [&](const ValueInfo& info, const BYTE* data, DWORD size) {
                                if (!reserved)
                                {
                                    if (result.info_valid)
                                    {
                                        snapshot.values.reserve(result.info.value_count);
                                    }
                                    reserved = true;
                                }
                                RegistryValue value;
                                value.name = info.name;
                                value.type = info.type;
                                if (data && size > 0)
                                {
                                    value.data.assign(data, data + size);
                                }
                                snapshot.values.push_back(std::move(value));
                                return !cancelled();
                            },
                            [&](const std::wstring& name) {
                                children.push_back(name);
                                return !cancelled();
                            },
                            MAXDWORD,
                            nullptr,
                            true,
                            link
                        ) &&
                        result.error == ERROR_SUCCESS && snapshot.complete;

    snapshot.class_name = std::move(result.options.class_name);
    snapshot.is_volatile = result.options.is_volatile;
    if (parts)
    {
        snapshot.last_write = result.info.last_write;
    }
    // partial snapshots are never restored, stop at the first unreadable part
    snapshot.children.reserve(children.size());
    for (const std::wstring& name : children)
    {
        if (!snapshot.complete)
        {
            return snapshot;
        }
        snapshot.children.push_back(Capture(registry_path::ChildNode(node, name), parts, cancel));
        snapshot.complete = snapshot.children.back().complete;
    }
    // a value or subkey deleted during the enumeration ends it early without an error
    if (!link && result.info_valid && (result.info.subkey_count != children.size() || result.info.value_count != snapshot.values.size()))
    {
        snapshot.complete = false;
    }
    return snapshot;
}

bool Restore(const RegistryNode& parent, const KeySnapshot& snapshot, bool* created);

bool Fill(const RegistryNode& node, const KeySnapshot& snapshot)
{
    for (const RegistryValue& value : snapshot.values)
    {
        if (!RegistryStore::SetValue(node, value.name, value.type, value.data))
        {
            return false;
        }
    }
    for (const KeySnapshot& child : snapshot.children)
    {
        if (!Restore(node, child, nullptr))
        {
            return false;
        }
    }
    return true;
}

bool WriteAttributes(const RegistryNode& node, const KeySnapshot& snapshot)
{
    // restore security last so a restrictive DACL cannot block other changes
    const bool timed = snapshot.last_write.dwLowDateTime || snapshot.last_write.dwHighDateTime;
    return (snapshot.security.empty() && !timed) ||
           RegistryStore::WriteKeySecurity(node, snapshot.security_parts, snapshot.security, timed ? &snapshot.last_write : nullptr);
}

bool Restore(const RegistryNode& parent, const KeySnapshot& snapshot, bool* created)
{
    if (snapshot.name.empty())
    {
        return false;
    }
    const RegistryNode node = registry_path::ChildNode(parent, snapshot.name);
    const KeyCreateOptions options = {snapshot.class_name, snapshot.is_volatile};
    const bool link = !snapshot.link_target.empty();
    if (!(link ? RegistryStore::CreateKeyLink(parent, snapshot.name, snapshot.link_target, options) : RegistryStore::CreateKey(parent, snapshot.name, options)))
    {
        return false;
    }
    if (created)
    {
        *created = true;
    }
    return (link || Fill(node, snapshot)) && WriteAttributes(node, snapshot);
}

// links stay unfollowed
bool ListChildren(const RegistryNode& node, std::vector<std::wstring>* children, bool* is_volatile)
{
    RegistryStore::KeyEnumResult result;
    result.want_options = is_volatile != nullptr;
    const bool enumerated = RegistryStore::EnumKeyStreaming(
        node,
        false,
        false,
        true,
        &result,
        {},
        [children](const std::wstring& name) {
            children->push_back(name);
            return true;
        },
        MAXDWORD,
        nullptr,
        false,
        true
    );
    if (is_volatile)
    {
        *is_volatile = result.options.is_volatile;
    }
    return enumerated && result.error == ERROR_SUCCESS;
}

// children of a volatile key are volatile too and ride in its snapshot
bool CollectVolatile(const RegistryNode& parent, const RegistryNode& node, std::vector<std::pair<RegistryNode, KeySnapshot>>* keys)
{
    std::vector<std::wstring> children;
    bool is_volatile = false;
    if (!ListChildren(node, &children, &is_volatile))
    {
        return false;
    }
    if (is_volatile)
    {
        keys->emplace_back(parent, CaptureKey(node));
        return keys->back().second.complete;
    }
    return std::all_of(children.begin(), children.end(), [&](const std::wstring& name) { return CollectVolatile(node, registry_path::ChildNode(node, name), keys); });
}

}

bool CaptureVolatileKeys(const RegistryNode& node, std::vector<std::pair<RegistryNode, KeySnapshot>>* keys)
{
    std::vector<std::wstring> children;
    return ListChildren(node, &children, nullptr) && std::all_of(children.begin(), children.end(), [&](const std::wstring& name) {
               return CollectVolatile(node, registry_path::ChildNode(node, name), keys);
           });
}

KeySnapshot CaptureKey(const RegistryNode& node, bool exact, const std::atomic_bool* cancel)
{
    const util::PrivilegeScope audit({SE_SECURITY_NAME});
    KeySnapshot snapshot = Capture(node, exact ? kExactSecurity | (audit.held() ? SACL_SECURITY_INFORMATION : 0) : 0, cancel);
    return snapshot.complete ? std::move(snapshot) : KeySnapshot{.complete = false};
}

bool RestoreKey(const RegistryNode& parent, const KeySnapshot& snapshot)
{
    // partial snapshots would be restored as complete keys
    if (!snapshot.complete)
    {
        return false;
    }
    const util::PrivilegeScope owner({SE_RESTORE_NAME});
    const util::PrivilegeScope audit({SE_SECURITY_NAME});
    bool created = false;
    if (Restore(parent, snapshot, &created))
    {
        return true;
    }
    if (created)
    {
        RegistryStore::DeleteKey(registry_path::ChildNode(parent, snapshot.name));
    }
    return false;
}

}
