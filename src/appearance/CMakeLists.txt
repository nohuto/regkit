regkit_add_library(regkit_appearance
    default_font.cpp
    dialog_layout.cpp
    feedback.cpp
    font_metrics.cpp
    font_picker.cpp
    gdi_cache.cpp
    icon_loader.cpp
    presets.cpp
    preset_editor.cpp
    theme.cpp
)
target_include_directories(regkit_appearance PUBLIC
    "${PROJECT_SOURCE_DIR}/src"
)
target_link_libraries(regkit_appearance PRIVATE
    regkit_win32
    advapi32
    comctl32
    comdlg32
    dwmapi
    shell32
    uxtheme
)
