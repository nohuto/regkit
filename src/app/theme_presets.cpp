// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "app/theme_presets.h"

#include <algorithm>
#include <cwctype>
#include <fstream>
#include <sstream>

#include "win32/win32_helpers.h"

namespace regkit {

namespace {

constexpr wchar_t kPresetSection[] = L"[preset]";

ThemePreset MakePreset(const wchar_t* name, const ThemeColors& colors, bool is_dark) {
  ThemePreset preset;
  preset.name = name ? name : L"";
  preset.colors = colors;
  preset.is_dark = is_dark;
  return preset;
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    out.push_back(static_cast<wchar_t>(towlower(ch)));
  }
  return out;
}

std::wstring Trim(const std::wstring& text) {
  size_t start = 0;
  while (start < text.size() && (text[start] == L' ' || text[start] == L'\t')) {
    ++start;
  }
  size_t end = text.size();
  while (end > start && (text[end - 1] == L' ' || text[end - 1] == L'\t')) {
    --end;
  }
  return text.substr(start, end - start);
}

void AddPresetIfMissing(std::vector<ThemePreset>* presets, const ThemePreset& preset) {
  if (!presets) {
    return;
  }
  auto it = std::find_if(presets->begin(), presets->end(), [&](const ThemePreset& existing) { return _wcsicmp(existing.name.c_str(), preset.name.c_str()) == 0; });
  if (it == presets->end()) {
    presets->push_back(preset);
  }
}

ThemeColors DarkDefaults() {
  ThemeColors colors;
  colors.background = RGB(20, 20, 20);
  colors.panel = RGB(20, 20, 20);
  colors.surface = RGB(34, 34, 34);
  colors.header = colors.surface;
  colors.border = RGB(66, 66, 66);
  colors.text = RGB(200, 200, 200);
  colors.muted_text = RGB(170, 170, 170);
  colors.accent = RGB(90, 162, 255);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(44, 44, 44);
  colors.focus = colors.accent;
  return colors;
}

ThemeColors LightDefaults() {
  ThemeColors colors;
  colors.background = RGB(245, 245, 245);
  colors.panel = RGB(255, 255, 255);
  colors.surface = RGB(242, 242, 242);
  colors.header = colors.surface;
  colors.border = RGB(204, 204, 204);
  colors.text = RGB(32, 32, 32);
  colors.muted_text = RGB(96, 96, 96);
  colors.accent = RGB(0, 120, 215);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(236, 236, 236);
  colors.focus = colors.accent;
  return colors;
}

ThemeColors SolarizedDark() {
  ThemeColors colors;
  colors.background = RGB(0, 43, 54);
  colors.panel = RGB(7, 54, 66);
  colors.surface = RGB(10, 60, 71);
  colors.header = colors.surface;
  colors.border = RGB(15, 59, 70);
  colors.text = RGB(147, 161, 161);
  colors.muted_text = RGB(131, 148, 150);
  colors.accent = RGB(181, 137, 0);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(10, 60, 71);
  colors.focus = RGB(181, 137, 0);
  return colors;
}

ThemeColors SolarizedLight() {
  ThemeColors colors;
  colors.background = RGB(253, 246, 227);
  colors.panel = RGB(238, 232, 213);
  colors.surface = RGB(228, 221, 200);
  colors.header = colors.surface;
  colors.border = RGB(214, 207, 181);
  colors.text = RGB(88, 110, 117);
  colors.muted_text = RGB(101, 123, 131);
  colors.accent = RGB(181, 137, 0);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(228, 221, 200);
  colors.focus = RGB(181, 137, 0);
  return colors;
}

ThemeColors NordDark() {
  ThemeColors colors;
  colors.background = RGB(46, 52, 64);
  colors.panel = RGB(59, 66, 82);
  colors.surface = RGB(67, 76, 94);
  colors.header = colors.surface;
  colors.border = RGB(76, 86, 106);
  colors.text = RGB(229, 233, 240);
  colors.muted_text = RGB(167, 177, 194);
  colors.accent = RGB(136, 192, 208);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(67, 76, 94);
  colors.focus = RGB(136, 192, 208);
  return colors;
}

ThemeColors Dracula() {
  ThemeColors colors;
  colors.background = RGB(40, 42, 54);
  colors.panel = RGB(52, 55, 70);
  colors.surface = RGB(59, 63, 82);
  colors.header = colors.surface;
  colors.border = RGB(68, 71, 90);
  colors.text = RGB(248, 248, 242);
  colors.muted_text = RGB(191, 191, 191);
  colors.accent = RGB(189, 147, 249);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(59, 63, 82);
  colors.focus = RGB(189, 147, 249);
  return colors;
}

ThemeColors GruvboxDark() {
  ThemeColors colors;
  colors.background = RGB(40, 40, 40);
  colors.panel = RGB(50, 48, 47);
  colors.surface = RGB(60, 56, 54);
  colors.header = colors.surface;
  colors.border = RGB(80, 73, 69);
  colors.text = RGB(235, 219, 178);
  colors.muted_text = RGB(189, 174, 147);
  colors.accent = RGB(250, 189, 47);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(60, 56, 54);
  colors.focus = RGB(250, 189, 47);
  return colors;
}

ThemeColors GruvboxLight() {
  ThemeColors colors;
  colors.background = RGB(251, 241, 199);
  colors.panel = RGB(242, 229, 188);
  colors.surface = RGB(235, 219, 178);
  colors.header = colors.surface;
  colors.border = RGB(213, 196, 161);
  colors.text = RGB(60, 56, 54);
  colors.muted_text = RGB(124, 111, 100);
  colors.accent = RGB(215, 153, 33);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(235, 219, 178);
  colors.focus = RGB(215, 153, 33);
  return colors;
}

ThemeColors CatppuccinMocha() {
  ThemeColors colors;
  colors.background = RGB(30, 30, 46);
  colors.panel = RGB(42, 43, 60);
  colors.surface = RGB(49, 50, 68);
  colors.header = colors.surface;
  colors.border = RGB(69, 71, 90);
  colors.text = RGB(205, 214, 244);
  colors.muted_text = RGB(166, 173, 200);
  colors.accent = RGB(137, 180, 250);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(49, 50, 68);
  colors.focus = RGB(137, 180, 250);
  return colors;
}

ThemeColors CatppuccinMacchiato() {
  ThemeColors colors;
  colors.background = RGB(36, 39, 58);
  colors.panel = RGB(48, 52, 70);
  colors.surface = RGB(54, 58, 79);
  colors.header = colors.surface;
  colors.border = RGB(73, 77, 100);
  colors.text = RGB(202, 211, 245);
  colors.muted_text = RGB(165, 173, 203);
  colors.accent = RGB(138, 173, 244);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(54, 58, 79);
  colors.focus = RGB(138, 173, 244);
  return colors;
}

ThemeColors CatppuccinFrappe() {
  ThemeColors colors;
  colors.background = RGB(48, 52, 70);
  colors.panel = RGB(65, 69, 89);
  colors.surface = RGB(81, 87, 109);
  colors.header = colors.surface;
  colors.border = RGB(98, 104, 128);
  colors.text = RGB(198, 208, 245);
  colors.muted_text = RGB(181, 191, 226);
  colors.accent = RGB(140, 170, 238);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(81, 87, 109);
  colors.focus = RGB(140, 170, 238);
  return colors;
}

ThemeColors CatppuccinLatte() {
  ThemeColors colors;
  colors.background = RGB(239, 241, 245);
  colors.panel = RGB(230, 233, 239);
  colors.surface = RGB(220, 224, 232);
  colors.header = colors.surface;
  colors.border = RGB(204, 208, 218);
  colors.text = RGB(76, 79, 105);
  colors.muted_text = RGB(108, 111, 133);
  colors.accent = RGB(30, 102, 245);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(220, 224, 232);
  colors.focus = RGB(30, 102, 245);
  return colors;
}

ThemeColors TokyoNight() {
  ThemeColors colors;
  colors.background = RGB(26, 27, 38);
  colors.panel = RGB(36, 40, 59);
  colors.surface = RGB(47, 51, 77);
  colors.header = colors.surface;
  colors.border = RGB(65, 72, 104);
  colors.text = RGB(192, 202, 245);
  colors.muted_text = RGB(169, 177, 214);
  colors.accent = RGB(122, 162, 247);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(47, 51, 77);
  colors.focus = RGB(122, 162, 247);
  return colors;
}

ThemeColors OneDark() {
  ThemeColors colors;
  colors.background = RGB(40, 44, 52);
  colors.panel = RGB(47, 52, 63);
  colors.surface = RGB(59, 64, 74);
  colors.header = colors.surface;
  colors.border = RGB(62, 68, 81);
  colors.text = RGB(171, 178, 191);
  colors.muted_text = RGB(139, 147, 165);
  colors.accent = RGB(97, 175, 239);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(59, 64, 74);
  colors.focus = RGB(97, 175, 239);
  return colors;
}

ThemeColors OneLight() {
  ThemeColors colors;
  colors.background = RGB(250, 250, 250);
  colors.panel = RGB(242, 242, 242);
  colors.surface = RGB(231, 231, 231);
  colors.header = colors.surface;
  colors.border = RGB(208, 208, 208);
  colors.text = RGB(56, 58, 66);
  colors.muted_text = RGB(107, 111, 119);
  colors.accent = RGB(64, 120, 242);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(231, 231, 231);
  colors.focus = RGB(64, 120, 242);
  return colors;
}

ThemeColors Monokai() {
  ThemeColors colors;
  colors.background = RGB(39, 40, 34);
  colors.panel = RGB(45, 46, 39);
  colors.surface = RGB(58, 59, 51);
  colors.header = colors.surface;
  colors.border = RGB(62, 61, 50);
  colors.text = RGB(248, 248, 242);
  colors.muted_text = RGB(197, 197, 190);
  colors.accent = RGB(166, 226, 46);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(58, 59, 51);
  colors.focus = RGB(166, 226, 46);
  return colors;
}

ThemeColors AyuDark() {
  ThemeColors colors;
  colors.background = RGB(15, 20, 25);
  colors.panel = RGB(21, 26, 33);
  colors.surface = RGB(27, 34, 43);
  colors.header = colors.surface;
  colors.border = RGB(37, 51, 64);
  colors.text = RGB(230, 225, 207);
  colors.muted_text = RGB(166, 179, 191);
  colors.accent = RGB(255, 180, 84);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(27, 34, 43);
  colors.focus = RGB(255, 180, 84);
  return colors;
}

ThemeColors AyuLight() {
  ThemeColors colors;
  colors.background = RGB(250, 250, 250);
  colors.panel = RGB(243, 243, 243);
  colors.surface = RGB(232, 232, 232);
  colors.header = colors.surface;
  colors.border = RGB(214, 214, 214);
  colors.text = RGB(92, 103, 115);
  colors.muted_text = RGB(138, 145, 153);
  colors.accent = RGB(242, 151, 24);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(232, 232, 232);
  colors.focus = RGB(242, 151, 24);
  return colors;
}

ThemeColors EverforestDark() {
  ThemeColors colors;
  colors.background = RGB(43, 51, 57);
  colors.panel = RGB(52, 63, 68);
  colors.surface = RGB(60, 71, 77);
  colors.header = colors.surface;
  colors.border = RGB(61, 72, 77);
  colors.text = RGB(211, 198, 170);
  colors.muted_text = RGB(157, 169, 160);
  colors.accent = RGB(167, 192, 128);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(60, 71, 77);
  colors.focus = RGB(167, 192, 128);
  return colors;
}

ThemeColors EverforestLight() {
  ThemeColors colors;
  colors.background = RGB(243, 234, 211);
  colors.panel = RGB(232, 223, 198);
  colors.surface = RGB(223, 212, 181);
  colors.header = colors.surface;
  colors.border = RGB(211, 198, 170);
  colors.text = RGB(92, 106, 114);
  colors.muted_text = RGB(127, 140, 141);
  colors.accent = RGB(141, 161, 1);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(223, 212, 181);
  colors.focus = RGB(141, 161, 1);
  return colors;
}

ThemeColors Material() {
  ThemeColors colors;
  colors.background = RGB(38, 50, 56);
  colors.panel = RGB(47, 59, 67);
  colors.surface = RGB(54, 69, 79);
  colors.header = colors.surface;
  colors.border = RGB(55, 71, 79);
  colors.text = RGB(207, 216, 220);
  colors.muted_text = RGB(176, 190, 197);
  colors.accent = RGB(128, 203, 196);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(54, 69, 79);
  colors.focus = RGB(128, 203, 196);
  return colors;
}

ThemeColors Horizon() {
  ThemeColors colors;
  colors.background = RGB(28, 30, 38);
  colors.panel = RGB(35, 37, 48);
  colors.surface = RGB(45, 47, 58);
  colors.header = colors.surface;
  colors.border = RGB(46, 48, 62);
  colors.text = RGB(224, 224, 224);
  colors.muted_text = RGB(157, 160, 162);
  colors.accent = RGB(233, 86, 120);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(45, 47, 58);
  colors.focus = RGB(233, 86, 120);
  return colors;
}

ThemeColors NightOwl() {
  ThemeColors colors;
  colors.background = RGB(1, 22, 39);
  colors.panel = RGB(11, 37, 58);
  colors.surface = RGB(17, 50, 77);
  colors.header = colors.surface;
  colors.border = RGB(18, 48, 71);
  colors.text = RGB(214, 222, 235);
  colors.muted_text = RGB(159, 179, 200);
  colors.accent = RGB(130, 170, 255);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(17, 50, 77);
  colors.focus = RGB(130, 170, 255);
  return colors;
}

ThemeColors RosePine() {
  ThemeColors colors;
  colors.background = RGB(25, 23, 36);
  colors.panel = RGB(31, 29, 46);
  colors.surface = RGB(38, 35, 58);
  colors.header = colors.surface;
  colors.border = RGB(64, 61, 82);
  colors.text = RGB(224, 222, 244);
  colors.muted_text = RGB(156, 154, 179);
  colors.accent = RGB(235, 111, 146);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(38, 35, 58);
  colors.focus = RGB(235, 111, 146);
  return colors;
}

ThemeColors RosePineMoon() {
  ThemeColors colors;
  colors.background = RGB(35, 33, 54);
  colors.panel = RGB(42, 39, 63);
  colors.surface = RGB(49, 48, 74);
  colors.header = colors.surface;
  colors.border = RGB(68, 65, 90);
  colors.text = RGB(224, 222, 244);
  colors.muted_text = RGB(179, 176, 214);
  colors.accent = RGB(234, 154, 151);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(49, 48, 74);
  colors.focus = RGB(234, 154, 151);
  return colors;
}

ThemeColors KanagawaWave() {
  ThemeColors colors;
  colors.background = RGB(31, 31, 40);
  colors.panel = RGB(42, 42, 55);
  colors.surface = RGB(54, 54, 70);
  colors.header = colors.surface;
  colors.border = RGB(59, 59, 79);
  colors.text = RGB(220, 215, 186);
  colors.muted_text = RGB(166, 166, 156);
  colors.accent = RGB(126, 156, 216);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(54, 54, 70);
  colors.focus = RGB(126, 156, 216);
  return colors;
}

ThemeColors KanagawaDragon() {
  ThemeColors colors;
  colors.background = RGB(24, 22, 22);
  colors.panel = RGB(31, 31, 31);
  colors.surface = RGB(38, 38, 38);
  colors.header = colors.surface;
  colors.border = RGB(45, 42, 46);
  colors.text = RGB(197, 201, 197);
  colors.muted_text = RGB(166, 166, 156);
  colors.accent = RGB(127, 180, 202);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(38, 38, 38);
  colors.focus = RGB(127, 180, 202);
  return colors;
}

ThemeColors KanagawaLotus() {
  ThemeColors colors;
  colors.background = RGB(242, 236, 188);
  colors.panel = RGB(231, 221, 176);
  colors.surface = RGB(223, 212, 164);
  colors.header = colors.surface;
  colors.border = RGB(200, 192, 160);
  colors.text = RGB(77, 74, 65);
  colors.muted_text = RGB(116, 108, 93);
  colors.accent = RGB(196, 109, 137);
  colors.selection = colors.panel;
  colors.selection_text = colors.text;
  colors.hover = RGB(223, 212, 164);
  colors.focus = RGB(196, 109, 137);
  return colors;
}

void WritePreset(std::wofstream& file, const ThemePreset& preset) {
  file << kPresetSection << L"\n";
  file << L"name=" << preset.name << L"\n";
  file << L"dark=" << (preset.is_dark ? L"1" : L"0") << L"\n";
  file << L"background=" << FormatColorHex(preset.colors.background) << L"\n";
  file << L"panel=" << FormatColorHex(preset.colors.panel) << L"\n";
  file << L"surface=" << FormatColorHex(preset.colors.surface) << L"\n";
  file << L"header=" << FormatColorHex(preset.colors.header) << L"\n";
  file << L"border=" << FormatColorHex(preset.colors.border) << L"\n";
  file << L"text=" << FormatColorHex(preset.colors.text) << L"\n";
  file << L"muted_text=" << FormatColorHex(preset.colors.muted_text) << L"\n";
  file << L"accent=" << FormatColorHex(preset.colors.accent) << L"\n";
  file << L"selection=" << FormatColorHex(preset.colors.selection) << L"\n";
  file << L"selection_text=" << FormatColorHex(preset.colors.selection_text) << L"\n";
  file << L"hover=" << FormatColorHex(preset.colors.hover) << L"\n";
  file << L"focus=" << FormatColorHex(preset.colors.focus) << L"\n";
  file << L"\n";
}

void ApplyField(ThemePreset* preset, const std::wstring& key, const std::wstring& value) {
  if (!preset) {
    return;
  }
  std::wstring key_lower = ToLower(key);
  if (key_lower == L"name") {
    preset->name = value;
    return;
  }
  if (key_lower == L"dark") {
    preset->is_dark = (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0);
    return;
  }
  COLORREF color = RGB(0, 0, 0);
  if (!ParseColorHex(value, &color)) {
    return;
  }
  if (key_lower == L"background") {
    preset->colors.background = color;
  } else if (key_lower == L"panel") {
    preset->colors.panel = color;
  } else if (key_lower == L"surface") {
    preset->colors.surface = color;
  } else if (key_lower == L"header") {
    preset->colors.header = color;
  } else if (key_lower == L"border") {
    preset->colors.border = color;
  } else if (key_lower == L"text") {
    preset->colors.text = color;
  } else if (key_lower == L"muted_text") {
    preset->colors.muted_text = color;
  } else if (key_lower == L"accent") {
    preset->colors.accent = color;
  } else if (key_lower == L"selection") {
    preset->colors.selection = color;
  } else if (key_lower == L"selection_text") {
    preset->colors.selection_text = color;
  } else if (key_lower == L"hover") {
    preset->colors.hover = color;
  } else if (key_lower == L"focus") {
    preset->colors.focus = color;
  }
}

std::vector<ThemePreset> LoadFromStream(std::wistream& file) {
  std::vector<ThemePreset> presets;
  ThemePreset current;
  bool in_preset = false;
  std::wstring line;
  while (std::getline(file, line)) {
    line = Trim(line);
    if (line.empty()) {
      continue;
    }
    if (line == kPresetSection) {
      if (in_preset && !current.name.empty()) {
        presets.push_back(current);
      }
      current = ThemePreset{};
      in_preset = true;
      continue;
    }
    if (!in_preset) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    ApplyField(&current, Trim(line.substr(0, sep)), Trim(line.substr(sep + 1)));
  }
  if (in_preset && !current.name.empty()) {
    presets.push_back(current);
  }
  return presets;
}

} // namespace

std::wstring ThemePresetStore::PresetsPath() {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"theme_presets.rktheme");
}

std::vector<ThemePreset> ThemePresetStore::BuiltInPresets() {
  std::vector<ThemePreset> presets;
  presets.push_back(MakePreset(L"Default Dark", DarkDefaults(), true));
  presets.push_back(MakePreset(L"Default Light", LightDefaults(), false));
  presets.push_back(MakePreset(L"Ayu Dark", AyuDark(), true));
  presets.push_back(MakePreset(L"Ayu Light", AyuLight(), false));
  presets.push_back(MakePreset(L"Catppuccin Frappe", CatppuccinFrappe(), true));
  presets.push_back(MakePreset(L"Catppuccin Latte", CatppuccinLatte(), false));
  presets.push_back(MakePreset(L"Catppuccin Macchiato", CatppuccinMacchiato(), true));
  presets.push_back(MakePreset(L"Catppuccin Mocha", CatppuccinMocha(), true));
  presets.push_back(MakePreset(L"Dracula", Dracula(), true));
  presets.push_back(MakePreset(L"Everforest Dark", EverforestDark(), true));
  presets.push_back(MakePreset(L"Everforest Light", EverforestLight(), false));
  presets.push_back(MakePreset(L"Gruvbox Dark", GruvboxDark(), true));
  presets.push_back(MakePreset(L"Gruvbox Light", GruvboxLight(), false));
  presets.push_back(MakePreset(L"Horizon", Horizon(), true));
  presets.push_back(MakePreset(L"Kanagawa Dragon", KanagawaDragon(), true));
  presets.push_back(MakePreset(L"Kanagawa Lotus", KanagawaLotus(), false));
  presets.push_back(MakePreset(L"Kanagawa Wave", KanagawaWave(), true));
  presets.push_back(MakePreset(L"Material", Material(), true));
  presets.push_back(MakePreset(L"Monokai", Monokai(), true));
  presets.push_back(MakePreset(L"Night Owl", NightOwl(), true));
  presets.push_back(MakePreset(L"Nord", NordDark(), true));
  presets.push_back(MakePreset(L"One Dark", OneDark(), true));
  presets.push_back(MakePreset(L"One Light", OneLight(), false));
  presets.push_back(MakePreset(L"Rose Pine", RosePine(), true));
  presets.push_back(MakePreset(L"Rose Pine Moon", RosePineMoon(), true));
  presets.push_back(MakePreset(L"Solarized Dark", SolarizedDark(), true));
  presets.push_back(MakePreset(L"Solarized Light", SolarizedLight(), false));
  presets.push_back(MakePreset(L"Tokyo Night", TokyoNight(), true));
  return presets;
}

bool ThemePresetStore::Load(std::vector<ThemePreset>* presets, std::wstring* error) {
  if (presets) {
    presets->clear();
  }
  if (error) {
    error->clear();
  }
  std::wstring path = PresetsPath();
  if (path.empty()) {
    if (error) {
      *error = L"Failed to resolve the theme presets path.";
    }
    return false;
  }
  std::wifstream file(path);
  if (!file.is_open()) {
    return false;
  }
  std::vector<ThemePreset> loaded = LoadFromStream(file);
  if (presets) {
    *presets = std::move(loaded);
  }
  return true;
}

bool ThemePresetStore::Save(const std::vector<ThemePreset>& presets, std::wstring* error) {
  if (error) {
    error->clear();
  }
  std::wstring path = PresetsPath();
  if (path.empty()) {
    if (error) {
      *error = L"Failed to resolve the theme presets path.";
    }
    return false;
  }
  std::wofstream file(path, std::ios::trunc);
  if (!file.is_open()) {
    if (error) {
      *error = L"Failed to open the theme presets file.";
    }
    return false;
  }
  for (const auto& preset : presets) {
    if (preset.name.empty()) {
      continue;
    }
    WritePreset(file, preset);
  }
  return true;
}

bool ThemePresetStore::ImportFromFile(const std::wstring& path, std::vector<ThemePreset>* presets, std::wstring* error) {
  if (presets) {
    presets->clear();
  }
  if (error) {
    error->clear();
  }
  if (path.empty()) {
    if (error) {
      *error = L"Invalid theme preset file path.";
    }
    return false;
  }
  std::wifstream file(path);
  if (!file.is_open()) {
    if (error) {
      *error = L"Failed to open the theme preset file.";
    }
    return false;
  }
  std::vector<ThemePreset> loaded = LoadFromStream(file);
  if (presets) {
    *presets = std::move(loaded);
  }
  return true;
}

bool ThemePresetStore::ExportToFile(const std::wstring& path, const std::vector<ThemePreset>& presets, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (path.empty()) {
    if (error) {
      *error = L"Invalid theme preset file path.";
    }
    return false;
  }
  std::wofstream file(path, std::ios::trunc);
  if (!file.is_open()) {
    if (error) {
      *error = L"Failed to write the theme preset file.";
    }
    return false;
  }
  for (const auto& preset : presets) {
    if (preset.name.empty()) {
      continue;
    }
    WritePreset(file, preset);
  }
  return true;
}

std::wstring FormatColorHex(COLORREF color) {
  wchar_t buffer[16] = {};
  swprintf_s(buffer, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
  return buffer;
}

bool ParseColorHex(const std::wstring& text, COLORREF* color) {
  if (!color) {
    return false;
  }
  std::wstring value = Trim(text);
  if (value.size() == 7 && value[0] == L'#') {
    unsigned int rgb = 0;
    if (swscanf_s(value.c_str() + 1, L"%06x", &rgb) == 1) {
      *color = RGB((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
      return true;
    }
  }
  if (value.rfind(L"0x", 0) == 0 && value.size() >= 8) {
    unsigned int rgb = 0;
    if (swscanf_s(value.c_str() + 2, L"%06x", &rgb) == 1) {
      *color = RGB((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
      return true;
    }
  }
  if (value.find(L',') != std::wstring::npos) {
    unsigned int r = 0;
    unsigned int g = 0;
    unsigned int b = 0;
    if (swscanf_s(value.c_str(), L"%u,%u,%u", &r, &g, &b) == 3) {
      if (r <= 255 && g <= 255 && b <= 255) {
        *color = RGB(r, g, b);
        return true;
      }
    }
  }
  return false;
}

} // namespace regkit
