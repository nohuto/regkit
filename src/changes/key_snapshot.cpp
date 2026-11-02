// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/key_snapshot.h"

#include "registry/registry_path.h"
#include "win32/process_rights.h"

namespace regkit::changes
{
namespace
{

constexpr SECURITY_INFORMATION kExactSecurity =
    OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION | LABEL_SECURITY_INFORMATION;

KeySnapshot Capture(const RegistryNode& node, SECURITY_INFORMATION parts)
{
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
                                return true;
                            },
                            [&](const std::wstring& name) {
                                children.push_back(name);
                                return true;
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
    snapshot.children.reserve(children.size());
    for (const std::wstring& name : children)
    {
        snapshot.children.push_back(Capture(registry_path::ChildNode(node, name), parts));
        if (!snapshot.children.back().complete)
        {
            snapshot.complete = false;
        }
    }
    if (!link && result.info_valid && result.info.subkey_count != children.size())
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

}

bool ReplaceKey(const RegistryNode& node, const KeySnapshot& snapshot)
{
    const util::PrivilegeScope owner({SE_RESTORE_NAME});
    const util::PrivilegeScope audit({SE_SECURITY_NAME});
    std::vector<std::wstring> values;
    std::vector<std::wstring> children;
    RegistryStore::KeyEnumResult result;
    if (!RegistryStore::EnumKeyStreaming(
            node,
            true,
            false,
            true,
            &result,
            [&](const ValueInfo& info, const BYTE*, DWORD) {
                values.push_back(info.name);
                return true;
            },
            [&](const std::wstring& name) {
                children.push_back(name);
                return true;
            }
        ) ||
        result.error != ERROR_SUCCESS)
    {
        return false;
    }
    for (const std::wstring& name : values)
    {
        if (!RegistryStore::DeleteValue(node, name))
        {
            return false;
        }
    }
    for (const std::wstring& name : children)
    {
        if (!RegistryStore::DeleteKey(registry_path::ChildNode(node, name)))
        {
            return false;
        }
    }
    return Fill(node, snapshot) && WriteAttributes(node, snapshot);
}

KeySnapshot CaptureKey(const RegistryNode& node, bool exact)
{
    const util::PrivilegeScope audit({SE_SECURITY_NAME});
    return Capture(node, exact ? kExactSecurity | (audit.held() ? SACL_SECURITY_INFORMATION : 0) : 0);
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
