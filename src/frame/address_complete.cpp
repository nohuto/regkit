// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

class RegistryAddressEnum : public ::IEnumString, public ::IACList {
public:
  RegistryAddressEnum(MainWindow::Impl* owner, HWND edit) : owner_(owner), edit_(edit) {}

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** out) override {
    if (!out) {
      return E_POINTER;
    }
    *out = nullptr;
    if (riid == IID_IUnknown || riid == IID_IEnumString) {
      *out = static_cast<::IEnumString*>(this);
      AddRef();
      return S_OK;
    }
    if (riid == IID_IACList) {
      *out = static_cast<::IACList*>(this);
      AddRef();
      return S_OK;
    }
    return E_NOINTERFACE;
  }

  ULONG STDMETHODCALLTYPE AddRef() override { return static_cast<ULONG>(InterlockedIncrement(&ref_count_)); }

  ULONG STDMETHODCALLTYPE Release() override {
    ULONG count = static_cast<ULONG>(InterlockedDecrement(&ref_count_));
    if (count == 0) {
      delete this;
    }
    return count;
  }

  HRESULT STDMETHODCALLTYPE Next(ULONG celt, LPOLESTR* rgelt, ULONG* pceltFetched) override {
    if (!rgelt) {
      return E_POINTER;
    }
    if (celt > 1 && !pceltFetched) {
      return E_POINTER;
    }
    UpdateSuggestionsIfNeeded();
    ULONG fetched = 0;
    for (; fetched < celt && index_ < suggestions_.size(); ++fetched, ++index_) {
      const std::wstring& item = suggestions_[index_];
      size_t bytes = (item.size() + 1) * sizeof(wchar_t);
      wchar_t* buffer = static_cast<wchar_t*>(CoTaskMemAlloc(bytes));
      if (!buffer) {
        for (ULONG i = 0; i < fetched; ++i) {
          CoTaskMemFree(rgelt[i]);
        }
        if (pceltFetched) {
          *pceltFetched = 0;
        }
        return E_OUTOFMEMORY;
      }
      wcscpy_s(buffer, item.size() + 1, item.c_str());
      rgelt[fetched] = buffer;
    }
    if (pceltFetched) {
      *pceltFetched = fetched;
    }
    return fetched == celt ? S_OK : S_FALSE;
  }

  HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override {
    UpdateSuggestionsIfNeeded();
    if (index_ + celt >= suggestions_.size()) {
      index_ = suggestions_.size();
      return S_FALSE;
    }
    index_ += celt;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Reset() override {
    UpdateSuggestionsIfNeeded();
    index_ = 0;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Clone(IEnumString** out) override {
    if (!out) {
      return E_POINTER;
    }
    auto* clone = new RegistryAddressEnum(owner_, edit_);
    clone->suggestions_ = suggestions_;
    clone->index_ = index_;
    clone->last_text_ = last_text_;
    clone->query_override_ = query_override_;
    *out = clone;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Expand(PCWSTR text) noexcept override {
    if (!text) {
      query_override_.clear();
      return S_OK;
    }
    query_override_ = text;
    suggestions_.clear();
    index_ = 0;
    last_text_.clear();
    return S_OK;
  }

private:
  std::wstring ReadEditText() const {
    if (!edit_) {
      return L"";
    }
    int length = GetWindowTextLengthW(edit_);
    if (length <= 0) {
      return L"";
    }
    std::wstring text(static_cast<size_t>(length) + 1, L'\0');
    GetWindowTextW(edit_, MutableData(text), length + 1);
    text.resize(static_cast<size_t>(length));
    return text;
  }

  void UpdateSuggestionsIfNeeded() {
    if (!owner_) {
      suggestions_.clear();
      index_ = 0;
      last_text_.clear();
      return;
    }
    std::wstring query = query_override_.empty() ? ReadEditText() : query_override_;
    if (query_override_.empty() && edit_ && !query.empty()) {
      DWORD sel_start = 0;
      DWORD sel_end = 0;
      SendMessageW(edit_, EM_GETSEL, reinterpret_cast<WPARAM>(&sel_start), reinterpret_cast<LPARAM>(&sel_end));
      if (sel_end > sel_start && sel_end == query.size()) {
        query = query.substr(0, sel_start);
      }
    }
    if (query == last_text_) {
      return;
    }
    last_text_ = query;
    suggestions_ = owner_->BuildAddressSuggestions(query);
    index_ = 0;
  }

  ~RegistryAddressEnum() = default;

  LONG ref_count_ = 1;
  MainWindow::Impl* owner_ = nullptr;
  HWND edit_ = nullptr;
  std::vector<std::wstring> suggestions_;
  size_t index_ = 0;
  std::wstring last_text_;
  std::wstring query_override_;
};

struct AutoCompleteThemeContext {
  HWND owner = nullptr;
  const Theme* theme = nullptr;
};

LRESULT CALLBACK AutoCompleteListBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId);
    break;
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK AutoCompletePopupSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  switch (msg) {
  case WM_NOTIFY: {
    auto* header = reinterpret_cast<NMHDR*>(lparam);
    if (header && header->code == NM_CUSTOMDRAW && WindowClassEquals(header->hwndFrom, WC_LISTVIEWW)) {
      auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
      const Theme& theme = Theme::Current();
      switch (draw->nmcd.dwDrawStage) {
      case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;
      case CDDS_ITEMPREPAINT: {
        if (draw->nmcd.uItemState & CDIS_SELECTED) {
          return CDRF_DODEFAULT;
        }
        COLORREF text = theme.TextColor();
        COLORREF background = theme.FieldColor();
        if (draw->nmcd.uItemState & CDIS_HOT) {
          background = theme.HoverColor();
        }
        draw->clrText = text;
        draw->clrTextBk = background;
        return CDRF_NEWFONT;
      }
      default:
        break;
      }
    }
    break;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().FieldBrush());
    return TRUE;
  }
  case WM_CTLCOLORLISTBOX:
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = CTLCOLOR_STATIC;
    if (msg == WM_CTLCOLOREDIT) {
      type = CTLCOLOR_EDIT;
    } else if (msg == WM_CTLCOLORLISTBOX) {
      type = CTLCOLOR_LISTBOX;
    }
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

BOOL CALLBACK ApplyAutoCompleteThemeProc(HWND hwnd, LPARAM lparam) {
  auto* ctx = reinterpret_cast<AutoCompleteThemeContext*>(lparam);
  if (!ctx || !ctx->theme) {
    return TRUE;
  }
  DWORD pid = 0;
  GetWindowThreadProcessId(hwnd, &pid);
  if (pid != GetCurrentProcessId()) {
    return TRUE;
  }

  bool is_dropdown = WindowClassEquals(hwnd, L"Auto-Suggest Dropdown") || WindowClassEquals(hwnd, L"Autocomplete") || WindowClassEquals(hwnd, L"AutoComplete");
  if (!is_dropdown) {
    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    if ((style & WS_POPUP) == 0) {
      return TRUE;
    }
    bool has_list_child = false;
    EnumChildWindows(
        hwnd,
        [](HWND child, LPARAM param) -> BOOL {
          auto* found = reinterpret_cast<bool*>(param);
          if (!found || *found) {
            return TRUE;
          }
          if (WindowClassEquals(child, WC_LISTVIEWW) || WindowClassEquals(child, WC_LISTBOXW)) {
            *found = true;
          }
          return TRUE;
        },
        reinterpret_cast<LPARAM>(&has_list_child));
    if (!has_list_child) {
      return TRUE;
    }
  }

  ctx->theme->ApplyToWindow(hwnd);
  if (!GetWindowSubclass(hwnd, AutoCompletePopupSubclassProc, kAutoCompletePopupSubclassId, nullptr)) {
    SetWindowSubclass(hwnd, AutoCompletePopupSubclassProc, kAutoCompletePopupSubclassId, 0);
  }
  EnumChildWindows(
      hwnd,
      [](HWND child, LPARAM param) -> BOOL {
        auto* theme = reinterpret_cast<const Theme*>(param);
        if (!theme) {
          return TRUE;
        }
        if (WindowClassEquals(child, WC_LISTVIEWW)) {
          theme->ApplyToListView(child);
        } else if (WindowClassEquals(child, WC_LISTBOXW) || WindowClassEquals(child, L"ComboLBox")) {
          AllowDarkModeForWindow(child, Theme::UseDarkMode());
          const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
          SetWindowTheme(child, theme_name, nullptr);
          if (!GetWindowSubclass(child, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId, nullptr)) {
            SetWindowSubclass(child, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId, 0);
          }
        } else {
          AllowDarkModeForWindow(child, Theme::UseDarkMode());
          const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
          SetWindowTheme(child, theme_name, nullptr);
        }
        return TRUE;
      },
      reinterpret_cast<LPARAM>(ctx->theme));
  InvalidateRect(hwnd, nullptr, TRUE);
  return TRUE;
}

void MainWindow::Impl::EnableAddressAutoComplete() {
  if (!browse_.address() || address_autocomplete_) {
    return;
  }
  ::IAutoComplete2* autocomplete = nullptr;
  HRESULT hr = CoCreateInstance(CLSID_AutoComplete, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&autocomplete));
  if (FAILED(hr) || !autocomplete) {
    return;
  }

  ::IEnumString* source = new RegistryAddressEnum(this, browse_.address());
  hr = autocomplete->Init(browse_.address(), source, nullptr, nullptr);
  if (FAILED(hr)) {
    source->Release();
    autocomplete->Release();
    return;
  }
  DWORD options = ACO_AUTOSUGGEST | ACO_AUTOAPPEND | ACO_UPDOWNKEYDROPSLIST | ACO_FILTERPREFIXES;
  autocomplete->SetOptions(options);
  address_autocomplete_ = autocomplete;
  address_autocomplete_source_ = source;
}

std::vector<std::wstring> MainWindow::Impl::BuildAddressSuggestions(const std::wstring& input) const {
  std::vector<std::wstring> items;
  std::wstring text = TrimWhitespace(input);
  for (auto& ch : text) {
    if (ch == L'/') {
      ch = L'\\';
    }
  }
  bool trailing_sep = !text.empty() && text.back() == L'\\';
  if (trailing_sep) {
    text.pop_back();
  }

  constexpr size_t kMaxSuggestions = 200;
  auto add_unique = [&](const std::wstring& value, std::unordered_set<std::wstring>* seen) {
    if (value.empty()) {
      return;
    }
    std::wstring key = ToLower(value);
    if (seen->insert(key).second) {
      items.push_back(value);
    }
  };

  size_t sep = text.find_last_of(L'\\');
  if (trailing_sep) {
    sep = text.size();
  }
  if (sep == std::wstring::npos) {
    std::unordered_set<std::wstring> seen;
    const std::wstring prefix = text;
    for (const auto& root : browse_.roots()) {
      if (prefix.empty() || StartsWithInsensitive(root.path_name, prefix)) {
        add_unique(root.path_name, &seen);
      }
    }
    struct RootAlias {
      const wchar_t* short_name;
      const wchar_t* full_name;
    };
    const RootAlias aliases[] = {
        {L"HKCR", L"HKEY_CLASSES_ROOT"},
        {L"HKCU", L"HKEY_CURRENT_USER"},
        {L"HKLM", L"HKEY_LOCAL_MACHINE"},
        {L"HKU", L"HKEY_USERS"},
        {L"HKCC", L"HKEY_CURRENT_CONFIG"},
    };
    for (const auto& alias : aliases) {
      if (prefix.empty() || StartsWithInsensitive(alias.short_name, prefix)) {
        add_unique(alias.short_name, &seen);
        add_unique(alias.full_name, &seen);
      }
    }
    if (items.size() > kMaxSuggestions) {
      items.resize(kMaxSuggestions);
    }
    return items;
  }

  std::wstring prefix = text.substr(0, sep);
  std::wstring partial;
  if (sep < text.size()) {
    partial = text.substr(sep + 1);
  }
  if (prefix.empty()) {
    prefix = text;
  }
  std::wstring normalized_prefix = NormalizeRegistryPath(prefix);
  std::wstring display_prefix = prefix;
  if (display_prefix.empty()) {
    display_prefix = normalized_prefix;
  }
  RegistryNode node;
  if (!ResolvePathToNode(normalized_prefix, &node)) {
    return items;
  }
  KeyInfo info = {};
  if (!RegistryStore::QueryKeyInfo(node, &info)) {
    return items;
  }
  auto subkeys = RegistryStore::EnumSubKeyNames(node, true);
  items.reserve(std::min(subkeys.size(), kMaxSuggestions));
  for (const auto& name : subkeys) {
    if (!partial.empty() && !StartsWithInsensitive(name, partial)) {
      continue;
    }
    items.push_back(display_prefix + L"\\" + name);
    if (items.size() >= kMaxSuggestions) {
      break;
    }
  }
  return items;
}

void MainWindow::Impl::ApplyAutoCompleteTheme() {
  if (!Theme::UseDarkMode()) {
    return;
  }
  AutoCompleteThemeContext ctx;
  ctx.owner = hwnd_;
  ctx.theme = &Theme::Current();
  EnumThreadWindows(GetCurrentThreadId(), ApplyAutoCompleteThemeProc, reinterpret_cast<LPARAM>(&ctx));
}

} // namespace regkit
