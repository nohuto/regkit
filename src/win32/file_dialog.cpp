// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/file_dialog.h"

#include "win32/handle_owner.h"
#include "win32/system_error.h"

#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>

#include <vector>

namespace regkit::win32 {

namespace {

template <typename T>
class ComPtr {
public:
  ComPtr() = default;
  ~ComPtr() {
    if (ptr_) {
      ptr_->Release();
    }
  }
  ComPtr(const ComPtr&) = delete;
  ComPtr& operator=(const ComPtr&) = delete;
  T** Receive() { return &ptr_; }
  T* operator->() const { return ptr_; }
  explicit operator bool() const { return ptr_ != nullptr; }

private:
  T* ptr_ = nullptr;
};

class CoTaskString {
public:
  ~CoTaskString() {
    if (text_) {
      CoTaskMemFree(text_);
    }
  }
  PWSTR* Receive() { return &text_; }
  PCWSTR Get() const { return text_; }

private:
  PWSTR text_ = nullptr;
};

std::vector<COMDLG_FILTERSPEC> ParseFilter(const wchar_t* filter) {
  std::vector<COMDLG_FILTERSPEC> specs;
  if (!filter) {
    return specs;
  }
  const wchar_t* cursor = filter;
  while (*cursor) {
    const wchar_t* name = cursor;
    cursor += wcslen(cursor) + 1;
    if (!*cursor) {
      break;
    }
    const wchar_t* spec = cursor;
    cursor += wcslen(cursor) + 1;
    specs.push_back({name, spec});
  }
  return specs;
}

HRESULT ShowDialog(HWND owner, REFCLSID clsid, const wchar_t* filter,
                   FILEOPENDIALOGOPTIONS extra_options,
                   const wchar_t* default_extension,
                   const wchar_t* suggested_name, std::wstring* path) {
  if (!path) {
    return E_POINTER;
  }
  ComPtr<IFileDialog> dialog;
  HRESULT hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
                                IID_PPV_ARGS(dialog.Receive()));
  if (FAILED(hr)) {
    return hr;
  }

  FILEOPENDIALOGOPTIONS options = 0;
  hr = dialog->GetOptions(&options);
  if (FAILED(hr)) {
    return hr;
  }
  hr = dialog->SetOptions(options | FOS_FORCEFILESYSTEM | extra_options);
  if (FAILED(hr)) {
    return hr;
  }

  const std::vector<COMDLG_FILTERSPEC> specs = ParseFilter(filter);
  if (!specs.empty()) {
    hr = dialog->SetFileTypes(static_cast<UINT>(specs.size()), specs.data());
    if (SUCCEEDED(hr)) {
      dialog->SetFileTypeIndex(1);
    }
  }
  if (default_extension && *default_extension) {
    dialog->SetDefaultExtension(default_extension);
  }
  if (suggested_name && *suggested_name) {
    dialog->SetFileName(suggested_name);
  }

  hr = dialog->Show(owner);
  if (FAILED(hr)) {
    return hr;
  }

  ComPtr<IShellItem> item;
  hr = dialog->GetResult(item.Receive());
  if (FAILED(hr)) {
    return hr;
  }
  CoTaskString text;
  hr = item->GetDisplayName(SIGDN_FILESYSPATH, text.Receive());
  if (FAILED(hr)) {
    return hr;
  }
  *path = text.Get() ? text.Get() : L"";
  return S_OK;
}

} // namespace

HRESULT ChooseFileToOpen(HWND owner, const wchar_t* filter, std::wstring* path) {
  return ShowDialog(owner, CLSID_FileOpenDialog, filter, 0, nullptr, nullptr, path);
}

HRESULT ChooseFileToSave(HWND owner, const wchar_t* filter,
                         const wchar_t* default_extension,
                         const wchar_t* suggested_name, std::wstring* path) {
  return ShowDialog(owner, CLSID_FileSaveDialog, filter, 0, default_extension,
                    suggested_name, path);
}

HRESULT ChooseFolder(HWND owner, std::wstring* path) {
  return ShowDialog(owner, CLSID_FileOpenDialog, nullptr, FOS_PICKFOLDERS,
                    nullptr, nullptr, path);
}

bool DialogCancelled(HRESULT hr) {
  return hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

HRESULT ShellOpen(HWND owner, const wchar_t* target) {
  if (!target || !*target) {
    return E_INVALIDARG;
  }
  SHELLEXECUTEINFOW info = {};
  info.cbSize = sizeof(info);
  info.fMask = SEE_MASK_NOASYNC | SEE_MASK_FLAG_NO_UI;
  info.hwnd = owner;
  info.lpVerb = L"open";
  info.lpFile = target;
  info.nShow = SW_SHOWNORMAL;
  if (!ShellExecuteExW(&info)) {
    return HRESULT_FROM_WIN32(GetLastError());
  }
  return S_OK;
}

HRESULT RevealInExplorer(const std::wstring& path) {
  if (path.empty()) {
    return E_INVALIDARG;
  }
  PIDLIST_ABSOLUTE item = nullptr;
  HRESULT hr = SHParseDisplayName(path.c_str(), nullptr, &item, 0, nullptr);
  if (FAILED(hr)) {
    return hr;
  }
  hr = SHOpenFolderAndSelectItems(item, 0, nullptr, 0);
  CoTaskMemFree(item);
  return hr;
}

std::wstring FormatDialogError(HRESULT hr) {
  if (HRESULT_FACILITY(hr) == FACILITY_WIN32) {
    return util::FormatWin32Error(static_cast<DWORD>(HRESULT_CODE(hr)));
  }
  wchar_t code[32] = {};
  swprintf_s(code, L"0x%08X", static_cast<unsigned>(hr));
  return std::wstring(L"The file dialog failed (") + code + L").";
}

} // namespace regkit::win32
